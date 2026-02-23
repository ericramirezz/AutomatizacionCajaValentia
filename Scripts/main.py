import open 
import tkinter as tk
import scipy.io

if __name__ == "__main__":
    root = tk.Tk()
    app = open.MatToPandasApp(root)
    root.mainloop()
    #print(df)