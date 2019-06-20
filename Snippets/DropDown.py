master =...
frame = ...

def change_dropdown(self, *args):
    name = self.tkvar2.get()
    print(name)

def create_dropdown(master, frame):
  pages = ['bill','bob','bunny']
  tkvar2 = tk.StringVar(master)
  tkvar2.trace('w', change_dropdown)
  tkvar2.set(pages[len(pages)-1]) # set the default option

  pagesPopup2 = tk.OptionMenu(frame, tkvar2, *pages)
  pagesPopup2.grid(row = 1, column =5, padx=10, pady=10, sticky=tk.EW)

create_dropdown(master, frame)
