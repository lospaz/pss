#!/usr/bin/env python
import codecs
import json
import os
import ssl
from os import path

import configparser
import tornado.web
import tornado.websocket
import tornado.httpserver
import tornado.ioloop
from win32com.client import Dispatch
from Tkinter import *

config = configparser.RawConfigParser()
config.read("config.cfg", encoding='utf-8')

class Printer:
    def __init__(self):
        self.label = self.getLabel()
        self.printer = self.initPrinter()

    def getLabel(self):
        curdir = None
        if getattr(sys, 'frozen', False):
            curdir = path.dirname(sys.executable)
        else:
            curdir = path.dirname(path.abspath(__file__))

        mylabel = path.join(curdir, 'library.label')
        if not path.isfile(mylabel):
            return 0
        return mylabel

    def initPrinter(self):
        try:
            labelCom = Dispatch('Dymo.DymoAddIn')
            isOpen = labelCom.Open(self.label)
            selectPrinter = 'DYMO LabelWriter 450'
            labelCom.SelectPrinter(selectPrinter)
            return labelCom
        except:
            print "Error during printing!"
            sys.exit(1)

    def printLabel(self, progressive, code):
        labelText = Dispatch('Dymo.DymoLabels')
        labelText.SetField('Progressivo', progressive)
        labelText.SetField('CODE', code)
        self.printer.StartPrintJob()
        self.printer.Print(1, False)
        self.printer.EndPrintJob()


class ChannelHandler(tornado.websocket.WebSocketHandler):

    @classmethod
    def urls(cls):
        return [
            (r'/', cls, {}),  # Route/Handler/kwargs
        ]

    def initialize(self):
        self.channel = None

    def open(self):
        print "opened"

    def on_message(self, message):
        try:
            data = json.loads(message)
            progressive = data['progressive']
            code = data['code']
            #printer.printLabel(progressive, code)
        except Exception as e:
            print(e)
            self.write_message("error")

    def on_close(self):
        print "closed"

    def check_origin(self, origin):
        return True


def initSocket(printer):
    # Create tornado application and supply URL routes
    app = tornado.web.Application(ChannelHandler.urls())
    # Setup HTTP Server
    ssl_ctx = ssl.create_default_context(ssl.Purpose.CLIENT_AUTH)
    ssl_ctx.load_cert_chain(
        os.path.join(str(config.get('CERTIFICATE', 'CRT'))),
        os.path.join(str(config.get('CERTIFICATE', 'KEY')))
    )
    http_server = tornado.httpserver.HTTPServer(app, ssl_options=ssl_ctx)
    http_server.listen(str(config.get('SERVER', 'PORT')), str(config.get('SERVER', 'ADDRESS')))
    # Start IO/Event loop
    tornado.ioloop.IOLoop.instance().start()


printer = Printer()
initSocket(printer)
