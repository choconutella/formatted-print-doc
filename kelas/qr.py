import qrcode
import os


class QrCode:
    def __init__(self, context):
        self.qr = qrcode.make(context)

    def save(self,name):
        self.qr.save(os.path.join('temp', name +'.jpg'))

    def delete(self,name):
        if os.path.exists(os.path.join('temp',name +'.jpg')):
            os.remove(os.path.join('temp',name +'.jpg'))


