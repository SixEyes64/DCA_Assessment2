import pyodbc
from datetime import datetime
import tkinter as tk
from sys import exit

'''
Author: Hunter, 2022001566
Pledge of Honour: I pledge by honour that this program is solely my work
Purpose: Python program that connects to an access database, providing backend and frontend (GUI) support
'''

# BACKEND STARTS HERE

# Template for displaying the database in the backend
formatTemp = '{0:<8}{1:<25}{2:<18}{3:<15}{4:<20}{5:10}{6:<7}'

def add_data():
    '''Adds entries to database'''
    # Connect python to MS Access using pyodbc
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=\\ebellfile\studenthome$\2022001566\4DCA\Assessments\Assessment 2\2022001566_dca_a2\Assessment2\2022001566_t2.accdb;')
    cursor = conn.cursor()

    # Get the input from the entry widgets
    MovieID = e1.get()
    movie_name = e2.get()
    rec_date = e3.get()
    Genre = e4.get()
    maturity_rating = e5.get()
    Stars = e6.get()
    Views = e7.get()

    # Finally insert the values into MS Access
    cursor.execute('INSERT INTO MovieList VALUES (?,?,?,?,?,?,?)', (MovieID, movie_name, rec_date, Genre, maturity_rating, Stars, Views))
    conn.commit()


def show_data():
# Connect Python to MS Access
    '''Output all rows in the terminal'''

    # Connect python to MS Access using pyodbc
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=\\ebellfile\studenthome$\2022001566\4DCA\Assessments\Assessment 2\2022001566_dca_a2\Assessment2\2022001566_t2.accdb;')
    cursor = conn.cursor()

    cursor.execute('select * from MovieList')

    # Prints the format
    print(formatTemp.format('ID', 'Movie Name', 'Release Date', 'Genre', 'Maturity Rating', 'Stars', 'Views'))
    print(f'{"_":_<105}') # Prints the set amount of underscores

    # Fetch all data in each row
    for row in cursor.fetchall():
        rec_date = datetime.strftime(row.release_date, '%d/%m/%Y') # The date in dd/mm/yyyy for each record

        # Print the rows 
        record = f'{row.MovieID:<8}{row.movie_name:<25}{rec_date:<18}{row.Genre:<15}{row.maturity_rating:<20}{row.Stars:<10}{row.Views:<7}' 
        print(record)


def quit_program():
    '''Exits the program, including the frontend and backend'''
    root.destroy()
    exit()


# GUI FRONTEND STARTS HERE
root = tk.Tk()

# Labels
tk.Label(root, text="ID").grid(row=0)
tk.Label(root, text="Movie Name").grid(row=1)
tk.Label(root, text="Release Date").grid(row=2)
tk.Label(root, text="Genre").grid(row=3)
tk.Label(root, text="Maturity Rating").grid(row=4)
tk.Label(root, text="Stars").grid(row=5)
tk.Label(root, text="Views").grid(row=6)

# Buttons
tk.Button(root, text="Show", command=show_data).grid(column=0,row=8)
tk.Button(root, text="Add", command=add_data).grid(column=1,row=8)
tk.Button(root, text="Exit", command=quit_program).grid(column=2,row=8)

# Text boxes (Entry)
e1 = tk.Entry(root)
e2 = tk.Entry(root)
e3 = tk.Entry(root)
e4 = tk.Entry(root)
e5 = tk.Entry(root)
e6 = tk.Entry(root)
e7 = tk.Entry(root)

# Placement of widgets
e1.grid(row=0, column=1)
e2.grid(row=1, column=1)
e3.grid(row=2, column=1)
e4.grid(row=3, column=1)
e5.grid(row=4, column=1)
e6.grid(row=5, column=1)
e7.grid(row=6, column=1)

# Runs the window
tk.mainloop()

