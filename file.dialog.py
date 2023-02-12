from tkinter import filedialog
import sys

def pick_file():
    try:
        args = sys.argv[(slice(1, 4))]
        action = filedialog.asksaveasfile
        if args[1].split('--type=')[1] == 'open':
            action = filedialog.askopenfile
        return action(title=args[0].split('--title=')[1], initialfile=args[2].split('--initial=')[1], filetypes=(('Excel file', '*.xlsx'), ('CSV file', '*.csv'))).name
    except Exception as e:
        return None

print(pick_file())