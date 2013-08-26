#!/usr/bin/env python
# -*- coding: UTF-8 -*-

'''
SAPI 5+ driver.

Copyright (c) 2009, 2013 Peter Parente

Permission to use, copy, modify, and distribute this software for any
purpose with or without fee is hereby granted, provided that the above
copyright notice and this permission notice appear in all copies.

THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
'''
#import comtypes.client
import win32com.client
import pythoncom
import time
import math
import weakref
from ..voice import Voice

# common voices
MSSAM = 'HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Speech\\Voices\\Tokens\\MSSam'
MSMARY = 'HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Speech\\Voices\\Tokens\\MSMary'
MSMIKE = 'HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Speech\\Voices\\Tokens\\MSMike'

# coeffs for wpm conversion
E_REG = {MSSAM : (137.89, 1.11),
         MSMARY : (156.63, 1.11),
         MSMIKE : (154.37, 1.11)}

def buildDriver(proxy):
    return SAPI5Driver(proxy)

class SAPI5Driver(object):
    def __init__(self, proxy):
        self._tts = win32com.client.Dispatch('SAPI.SPVoice')
        #self._tts = comtypes.client.CreateObject('SAPI.SPVoice')
        # all events
        self._tts.EventInterests = 33790
        self._advise = win32com.client.WithEvents(self._tts,
            SAPI5DriverEventSink)
        self._advise.setDriver(weakref.proxy(self))
        #self._debug = comtypes.client.ShowEvents(self._tts)
        #self._advise = comtypes.client.GetEvents(self._tts, self)
        self._proxy = proxy
        self._looping = False
        self._speaking = False
        self._stopping = False
        # initial rate
        self._rateWpm = 200
        self.setProperty('voice', self.getProperty('voice'))

        # 记录文本长度: 主要是录音时决定录音循环时间等
        self._textLength = 0


    def destroy(self):
        self._tts.EventInterests = 0

    def say(self, text):
        self._proxy.setBusy(True)
        self._proxy.notify('started-utterance')
        self._speaking = True
        self._tts.Speak(unicode(text), 19)


    ###HICK hack for rec
    def rec(self, text, out):
        self._textLength = len(text)

        self._proxy.setBusy(True)
        self._proxy.notify('started-utterance')
        self._speaking = True

        ##### 以下区别于 say 用于记录语音成文件
        self._stream = win32com.client.Dispatch('SAPI.SpFileStream')
        if out == None:
            out = "out.wav"
        self._stream.Open(out, 3) 
        self._tts.AudioOutputStream = self._stream

        self._tts.Speak(unicode(text), 19)

    def close(self):
        self._stream.Close()



    def stop(self):
        if not self._speaking:
            return
        self._proxy.setBusy(True)
        self._stopping = True
        self._tts.Speak('', 3)

    def _toVoice(self, attr):
        return Voice(attr.Id, attr.GetDescription())

    def _tokenFromId(self, id):
        tokens = self._tts.GetVoices()
        for token in tokens:
            if token.Id == id: return token
        raise ValueError('unknown voice id %s', id)

    def getProperty(self, name):
        if name == 'voices':
            return [self._toVoice(attr) for attr in self._tts.GetVoices()]
        elif name == 'voice':
            return self._tts.Voice.Id
        elif name == 'rate':
            return self._rateWpm
        elif name == 'volume':
            return self._tts.Volume/100.0
        else:
            raise KeyError('unknown property %s' % name)

    def setProperty(self, name, value):
        if name == 'voice':
            token = self._tokenFromId(value)
            self._tts.Voice = token
            a, b = E_REG.get(value, E_REG[MSMARY])
            self._tts.Rate = int(math.log(self._rateWpm/a, b))
        elif name == 'rate':
            id = self._tts.Voice.Id
            a, b = E_REG.get(id, E_REG[MSMARY])
            try:
                self._tts.Rate = int(math.log(value/a, b))
            except TypeError, e:
                raise ValueError(str(e))
            self._rateWpm = value
        elif name == 'volume':
            try:
                self._tts.Volume = int(round(value*100, 2))
            except TypeError, e:
                raise ValueError(str(e))
        else:
            raise KeyError('unknown property %s' % name)

    def startLoop(self):


        first = True
        self._looping = True
        while self._looping:
            if first:
                self._proxy.setBusy(False)
                first = False
            ###HICK 问题的症结所在了，自打设置了 filestream 这里就没完没了了， 不过仍然没看懂这个循环是怎么结束的?
            ###HICK 难道是执行到一定程度会抛出异常?  貌似是靠下面的 iterate ，在 driver.py 里有不断的调用它
            pythoncom.PumpWaitingMessages()
            time.sleep(0.05)


    def startLoopRec(self):

        ###HICK hack for rec
        i = 0

        ### 实际测试， startLoopRec 先于 rec 函数被调用。
        first = True
        self._looping = True
        while self._looping:
            if first:
                self._proxy.setBusy(False)
                first = False
            ###HICK 问题的症结所在了，自打设置了 filestream 这里就没完没了了， 不过仍然没看懂这个循环是怎么结束的?
            ###HICK 难道是执行到一定程度会抛出异常?  貌似是靠下面的 iterate ，在 driver.py 里有不断的调用它
            pythoncom.PumpWaitingMessages()



            ###HICK 经过测试，这里纯粹就是等待 tts 输出，调节目前 1s 比较合适。前面的 sleep 0.05秒，则合理 i == 20
            ###HICK try to exit just after several loops
            ### 实测取值跟文本长度相关，当然有些格式比如时间也不简单的是跟文本相关的
            time.sleep(0.05)
            i += 1
            ### 特别注意一开始 self._textLength 的值可能是 0
            ### 经验值，根据字符长度推算时间: len(str) / 8
            loop_times = 0 # 给个默认值，
            if self._textLength > 0:
                loop_times = self._textLength / 5
                # 不能太小了
                if loop_times < 10:
                    loop_times = 10

            
            if i > loop_times:
                self.endLoop()


    def endLoop(self):
        self._looping = False

    def iterate(self):
        self._proxy.setBusy(False)
        while 1:
            pythoncom.PumpWaitingMessages()
            yield

class SAPI5DriverEventSink(object):
    def __init__(self):
        self._driver = None

    def setDriver(self, driver):
        self._driver = driver

    def OnWord(self, stream, pos, char, length):
        self._driver._proxy.notify('started-word', location=char, length=length)

    def OnEndStream(self, stream, pos):
        d = self._driver
        if d._speaking:
            d._proxy.notify('finished-utterance', completed=not d._stopping)
        d._speaking = False
        d._stopping = False
        d._proxy.setBusy(False)