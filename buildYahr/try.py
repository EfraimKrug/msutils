import os
# [...]
scriptDir = os.path.dirname(os.path.realpath(__file__))
self.setWindowIcon(QtGui.QIcon(scriptDir + os.path.sep + 'logo.png'))
