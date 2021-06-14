# LB planner
This was my first coding project ever, after only having taken introductory coding modules (CS1010X) and after reading the docs for the different libraries. The program was written in Python 3, using modules like tkinter, openpyxl, pickle and binpacking. 
I created this program to give some structure to the grouping system of our weekly visits because before this, groupings were mostly random, and this made attendance hard to track. 
The main function of this program is to give suggestions to the user on which group each volunteer should be assigned to, based on the number of times the volunteer has previously visited the elderly in that group. Ultimately, the program merely suggests, and the final grouping decision lies with the user. The hope is that using data, volunteers can be more meaningfully allocated into groups.

## For those coming from the link in the planner user guide:
I guess if you see this you're interested in coding!

These are the files that I've used to create the app. The main file is app.py and the classes defined are in planner.py and people.py.
init.py is the file that I've used to create the first planner_data.pkl files, by reading the data, again, from excel files (though they have different formats from the ones i've detailed.)
If you're looking to manually input data use init.py, though you will need to change the execution codes at the bottom a little.
