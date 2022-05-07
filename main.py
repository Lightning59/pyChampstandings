""" main file - Run this should prompt user for a properly formatted Excel
file then update that file with pretty output of championship standings"""

import tkinter
import tkinter.filedialog


def get_file_path():
    rootwindow = tkinter.Tk()
    rootwindow.withdraw()

    file_path = tkinter.filedialog.askopenfilename()
    try:
        assert file_path != ''
        return file_path
    except AssertionError:
        print("You did not select a valid file.")
        raise FileNotFoundError


if __name__ == '__main__':
    file_path = get_file_path()
    print("about to print path")
    print(file_path)
