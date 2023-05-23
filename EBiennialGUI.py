import tkinter as tk

class Ebiennial_GUI:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Ebiennial Processing")
        self.input1_label = tk.Label(self.window, text="Transaction_ID")
        self.input1_label.grid(row=0,column=0)
        self.input1 = tk.Entry(self.window)
        self.input1.grid(row=0,column=1)

        self.URLQueryOutputlabel = tk.Label(self.window, text="Query 1")
        self.URLQueryOutputlabel.grid(row=1,column=0)
        self.URLQueryOutput = tk.Label(self.window,text="")
        self.URLQueryOutput.grid(row=1,column=1)

        self.Generate_query1_button= tk.Button(self.window, text="Generate query1",command=self.generate_query1)
        self.Generate_query1_button.grid(row=2, column=0)

        self.input2_label = tk.Label(self.window, text="URL")
        self.input2_label.grid(row=4,column=0)
        self.input2 = tk.Entry(self.window)
        self.input2.grid(row=4,column=1)

        self.get_DOS_ID_button= tk.Button(self.window, text="get DOS ID",command=self.generate_query1)
        self.Generate_query1_button.grid(row=5)

        self.input3_label = tk.Label(self.window, text="Invoice ID")
        self.input3_label.grid(row=6,column=0)
        self.input3 = tk.Entry(self.window)
        self.input3.grid(row=6,column=1)
    
    def generate_query1(self):
        pass

    def run(self):
        self.window.mainloop()

gui = Ebiennial_GUI()
gui.run()

