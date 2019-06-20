

def doStuff(files, workingFile):
  for f in files:
    print(f + ": " + workingFile)

def newButton(frame):
  files = ['filename', 'filename2']
  workingFile = 'myFile.txt'

  button01 = tk.Button(frame, text="Shift", command=partial(doStuff, files, workingFile))
  button01.grid(row=1, column=2, columnspan=1, padx=10, pady=10, sticky=tk.EW)
