# main file - Run this should prompt user for a properly formatted excel file then update that file with pretty output of championship standings







if __name__== '__main__':

    import tkinter as tk
    import tkinter.filedialog

    rootwindow=tk.Tk()
    rootwindow.withdraw()

    file_path=tkinter.filedialog.askopenfilenames()
    try:
        assert file_path!=''
    except:
        print("You did not select a valid file.")
        raise FileNotFoundError

    print(file_path)
    print(file_path)