class SearchNames:
    def __init__(self, var, data):
        self.var = var
        self.targetArray = []
        self.targetPointer = 0
        self.targetName = ""
        self.responseLabel = ""
        self.outLabel = ""
        self.data = data

    def initialize(self):
        self.targetArray = []
        self.targetPointer = 0

    def setOutLabel(self, str):
        self.outLabel = str

    def getOutLabel(self):
        return self.outLabel

    def getLengthTargetArray(self):
        return len(self.targetArray)

    def getTargetEntry(self):
        self.targetEntry = self.targetArray[self.targetPointer][1]
        return self.targetEntry

    def getTargetName(self):
        if len(self.targetArray) < 1 or len(self.targetArray[self.targetPointer]) < 1:
            return ""

        self.targetName = self.targetArray[self.targetPointer][0]
        return self.targetName

    def getNewEntry(self):
        #print(len(self.data))

        self.emptyEntry = self.data['ENTRIES'][0]
        for e in self.emptyEntry:
            self.emptyEntry[e] = ""
        self.emptyEntry["PayLevel"] = "1"
        return self.emptyEntry

    def saveNewEntry(self, newEntry):
        self.data.append(newEntry)
        #print(len(self.data))
        #print(self.data)

    def searchName(self, str):
        self.initialize()
        for entry in self.data['ENTRIES']:
            nArray = entry["Name"].split()
            for eName in nArray:
                if eName.lower().find(str.lower()) == 0:
                    self.targetArray.append([entry["Name"], entry])

        #print(self.targetArray[0][0])
        if len(self.targetArray) > 0 and len(self.targetArray[0]) > 0:
            self.setOutLabel(self.targetArray[0][0])
        else:
            self.setOutLabel("")

    def getEachStroke(self, key):
        if key.char.lower() == str('\b') and len(self.responseLabel) > 0:
            self.responseLabel = self.responseLabel[0:len(self.responseLabel)-1]

        if key.char.lower() == str(' '):
            self.responseLabel += key.char

        if key.char.lower() >= str('a') and key.char.lower() <= str('z'):
            self.responseLabel += key.char

        self.searchName(self.responseLabel)

    def getResponseLabel(self):
        return self.responseLabel

    def downArrow(self, key, setFunction):
        count = len(self.targetArray)
        if count < 2:
            return
        if self.targetPointer > count - 2:
            return
        #print("Incrementing")
        self.targetPointer += 1
        #self.setOutlabel()
        setFunction(self.getTargetName())

    def upArrow(self, key, setFunction):
        count = len(self.targetArray)
        if count < 2:
            return
        if self.targetPointer < 1:
            return
        #print("Decrementing")
        self.targetPointer -= 1
        #self.setOutlabel()
        setFunction(self.getTargetName())
