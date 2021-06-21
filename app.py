from tkinter import *
from tkinter import messagebox, filedialog
from planner import *
from datetime import datetime
from random import randint
import store
import os
from tkcalendar import DateEntry

# Absolutely necessary files to pass down: ./Data(both pkl), ./icon/LB.ico

LB = store.retrieve_pkl(os.path.join('Data', 'planner_data.pkl'))
SUNDAY = store.retrieve_pkl(os.path.join('Data', 'session_data.pkl'))
root = Tk()
root.title('LB planner')
root.iconbitmap(os.path.join('icon', 'LB.ico'))
root.geometry = '2194x1234'
root.resizable(False, False)
status_frame = Frame(root, bd=1, bg='white', relief='solid')
status_frame.pack()
status = Label(status_frame, text='Welcome! Planner ready.', padx=5, pady=5,
               font='Helvetica, 16', width=35, height=2, wraplength=400)
status.pack()

option_frame = Frame(root, bg='blue')
option_frame.pack(fill='both', expand=True)


def save():
    global LB, SUNDAY
    newpath = 'Data'
    if not os.path.exists(newpath):
        messagebox.showerror('LB planner', 'Data file is missing')
        return
    store.save_pkl(LB, os.path.join('Data', 'planner_data.pkl'))
    store.save_pkl(SUNDAY, os.path.join('Data', 'session_data.pkl'))


def create_session():
    global LB, SUNDAY
    if SUNDAY != None:
        resp1 = messagebox.askyesnocancel(
            title='LB planner', message='There seems to already be an exisiting planned session. Reprint detailing instead?')
        if resp1:
            SUNDAY.write_detailing('Sunday detailing.xlsx')
            status.config(
                text=f'Session created and detailing successfully exported.')
            return
        elif resp1 == None:
            return
        else:
            resp2 = messagebox.askyesno(
                title='LB planner', message='Reset detailing?')
            if resp2:
                SUNDAY = None
                create_session()
                return
            else:
                return
    top = Toplevel()
    top.resizable(False, False)
    top.title('Session config')
    top.geometry('250x175')
    F1 = Frame(top)
    F1.pack(pady=10)
    L1 = Label(F1, text='Choose date:')
    L1.pack(pady=5)
    cal = DateEntry(F1, showweeknumbers=False, showothermonthdays=False,
                    firstweekday='sunday', date_pattern='dd/mm/yy', borderwidth=2,)
    cal.pack()
    F2 = Frame(top)
    F2.pack()
    L2 = Label(F2, text='Elderly grouping file:')
    L2.pack(side='left')
    filename = ''

    def open():
        nonlocal filename
        filename = filedialog.askopenfilename(
            initialdir='.', title='Select file', filetypes=[("Microsoft Excel Worksheet", "*.xlsx")])
        L3 = Label(top, text=filename, wraplength=250)
        L3.pack(pady=5)
        top.lift()
    B1 = Button(F2, text='Select file', command=open)
    B1.pack(side='right')

    def exit():
        global SUNDAY
        nonlocal filename
        if not '.xlsx' in filename:
            messagebox.showerror(
                'LB planner', 'Invalid elderly grouping file.')
            return
        date = cal.get_date()
        date = datetime(date.year, date.month, date.day)
        try:
            SUNDAY = LB.create_session(date, filename)
        except KeyError as e:
            messagebox.showerror('LB planner', e)
            return
        SUNDAY.write_detailing('Sunday detailing.xlsx')
        status.config(
            text=f'Session created and detailing successfully exported.')
        top.destroy()

    B2 = Button(top, text='Confirm', command=exit)
    B2.pack(side='bottom')


button_1 = Button(option_frame, text='Create session',
                  command=create_session, width=20, height=2)
button_1.grid(column=1, row=1, ipady=5, pady=10, padx=5)


def pull_out():
    global LB, SUNDAY
    if not isinstance(SUNDAY, Session):
        messagebox.showerror('LB planner', 'No session detected.')
        return
    lst = list(map(lambda x: x.get_name(), SUNDAY.available_volunteers))
    top = Toplevel()
    top.resizable(False, False)
    top.title('Select person')
    top.geometry('250x175')
    option = StringVar(top)
    option.set('Choose person')
    choose = OptionMenu(top, option, *lst)
    choose.place(relx=0.5, rely=0.5, anchor='center')

    def end():
        name = option.get()
        if name == 'Choose person':
            messagebox.showerror('LB planner', 'No person selected.')
            return
        SUNDAY.pull_out(name)
        date = SUNDAY.date.date().strftime('%d-%m-%Y')
        status.config(
            text=f'Successfully pulled out {name} from {date} session')
        top.destroy()
    end = Button(top, text='Select', command=end)
    end.pack(side='bottom')


button_2 = Button(option_frame, text='Pull out',
                  command=pull_out, width=20, height=2)
button_2.grid(column=2, row=1, ipady=5, pady=10, padx=5)


def complete():
    global LB, SUNDAY
    if not isinstance(SUNDAY, Session):
        messagebox.showerror('LB planner', 'No session detected.')
        return
    date = SUNDAY.date.date().strftime('%d-%m-%Y')
    resp1 = messagebox.askyesno(
        'LB planner', f'Confirm complete of {date} Session?')
    if resp1:
        try:
            SUNDAY.complete()
            LB.update_from_session(SUNDAY)
        except KeyError as e:
            messagebox.showerror('LB planner', e)
            return
        SUNDAY = None
        status.config(
            text=f'{date} Session completed. Data archived and planner reset.')


button_3 = Button(option_frame, text='Complete session',
                  command=complete, width=20, height=2)
button_3.grid(column=3, row=1, ipady=5, pady=10, padx=5)


def cancel_session():
    global LB, SUNDAY
    if not isinstance(SUNDAY, Session):
        messagebox.showerror('LB planner', 'No session detected.')
        return
    date = SUNDAY.date.date().strftime('%d-%m-%Y')
    response = messagebox.askyesno(
        'LB planner', f'Confirm the cancellation of {date} session?')
    if response:
        LB.cancel_session(SUNDAY.date)
        SUNDAY = None
        status.config(text=f'{date} session successfully cancelled.')


button_4 = Button(option_frame, text='Cancel session',
                  command=cancel_session, width=20, height=2)
button_4.grid(column=4, row=1, ipady=5, pady=10, padx=5)


def output_avail_summ():
    global LB
    response = messagebox.askyesno(
        'LB planner', f'Output a summary of the available dates for all volunteers?')
    if response:
        LB.output_availability_summary(
            os.path.join('Output', str(datetime.today().date()) + ' availability summary.xlsx'))
        status.config(text=f'Successfully output availability summary.')


button_5 = Button(option_frame, text='Output availability\nsummary',
                  command=output_avail_summ, width=20, height=2)
button_5.grid(column=1, row=3, ipady=5, pady=10, padx=5)


def un_pull_out():
    global LB, SUNDAY
    if not isinstance(SUNDAY, Session):
        messagebox.showerror('LB planner', 'No session detected.')
        return
    if not SUNDAY.pulled_out:
        messagebox.showerror('LB planner', 'Nobody to pull out.')
        return
    lst = SUNDAY.pulled_out
    top = Toplevel()
    top.resizable(False, False)
    top.title('Select person')
    top.geometry('250x175')
    option = StringVar(top)
    option.set('Choose person')
    choose = OptionMenu(top, option, *lst)
    choose.place(relx=0.5, rely=0.5, anchor='center')

    def end():
        name = option.get()
        if name == 'Choose person':
            messagebox.showerror('LB planner', 'No person selected.')
            return
        SUNDAY.un_pull_out(name)
        date = SUNDAY.date.date().strftime('%d-%m-%Y')
        status.config(
            text=f'Successfully undid pull out of {name} from {date} session')
        top.destroy()
    end = Button(top, text='Select', command=end)
    end.pack(side='bottom')


button_6 = Button(option_frame, text='Undo pull out',
                  command=un_pull_out, width=20, height=2)
button_6.grid(column=2, row=2, ipady=5, pady=10, padx=5)


def update_from_elderlylist():
    global LB
    top = Toplevel()
    top.title('Choose file')
    top.geometry('250x175')
    top.resizable(False, False)

    def open():
        filename = filedialog.askopenfilename(
            initialdir='.', title='Select file', filetypes=[("Microsoft Excel Worksheet", "*.xlsx")])
        try:
            LB.update_from_elderlylist(filename)
        except ValueError as e:
            messagebox.showerror('LB planner', e)
            return
        status.config(text=f'Update from {filename} complete.')
        top.destroy()
    B1 = Button(top, text='Select file', command=open)
    B1.place(relx=0.5, rely=0.5, anchor='center')


button_7 = Button(option_frame, text='Update from elderly list',
                  command=update_from_elderlylist, width=20, height=2)
button_7.grid(column=3, row=2, ipady=5, pady=10, padx=5)


def update_from_volnlist():
    global LB
    top = Toplevel()
    top.title('Choose file')
    top.geometry('250x175')
    top.resizable(False, False)

    def open():
        filename = filedialog.askopenfilename(
            initialdir='.', title='Select file', filetypes=[("Microsoft Excel Worksheet", "*.xlsx")])
        if not os.path.exists(os.path.join('Archive', 'Previous volunteers')):
            os.makedirs(os.path.join('Archive', 'Previous volunteers'))
        try:
            LB.update_from_volnlist(filename)
        except ValueError as e:
            messagebox.showerror('LB planner', e)
            return
        status.config(text=f'Update from {filename} complete.')
        top.destroy()
    B1 = Button(top, text='Select file', command=open)
    B1.place(relx=0.5, rely=0.5, anchor='center')


button_8 = Button(option_frame, text='Update from volunteer list',
                  command=update_from_volnlist, width=20, height=2)
button_8.grid(column=4, row=2, ipady=5, pady=10, padx=5)


def output_eldgrp():
    global LB
    top = Toplevel()
    top.title('Write config')
    top.geometry('250x175')
    top.resizable(False, False)
    F1 = Frame(top)
    F1.place(relx=0.42, rely=0.25, anchor='center')
    L1 = Label(F1, text="# of Groups:", padx=3)
    L1.pack(side='left')
    E1 = Entry(F1, bd=2)
    E1.pack(side='left')
    F2 = Frame(top)
    F2.place(relx=0.5, rely=0.6, anchor='center')
    L2 = Label(F2, text="File name:", padx=3)
    L2.pack(side='left')
    L2a = Label(F2, text='.xlsx')
    L2a.pack(side='right')
    E2 = Entry(F2, bd=2)
    E2.pack(side='right')

    def end():
        try:
            n = int(E1.get())
        except ValueError:
            messagebox.showerror('LB planner', 'Invalid group size')
            return
        name = E2.get()
        if not os.path.exists('Elderly grouping'):
            os.makedirs('Elderly grouping')
        LB.write_elderly_grouping(n, os.path.join(
            'Elderly grouping', str(name) + '.xlsx'))
        status.config(
            text=f'Successfully output elderly grouping to {name}.xlsx')
        top.destroy()
    end = Button(top, text='Confirm', command=end)
    end.pack(side='bottom')


button_9 = Button(option_frame, text='Output elderly grouping', command=output_eldgrp,
                  width=20, height=2)
button_9.grid(column=2, row=3, ipady=5, pady=10, padx=5)


def output_volunteer_det():
    global LB, STATUS
    lst = ['All'] + list(map(lambda x: x.get_name(), LB.all_volunteers))
    top = Toplevel()
    top.resizable(False, False)
    top.title('Select person')
    top.geometry('250x175')
    option = StringVar(top)
    option.set('Choose person')
    choose = OptionMenu(top, option, *lst)
    choose.place(relx=0.5, rely=0.45, anchor='center')
    j = IntVar(top)
    c = Checkbutton(top, text='Show all', variable=j)
    c.place(relx=0.5, rely=0.7, anchor='center')

    def end():
        nonlocal j, option
        name = option.get()
        if name == 'Choose person':
            messagebox.showerror('LB planner', 'No person selected.')
            return
        if name == 'All':
            newpath = os.path.join('Output', str(
                datetime.today().date()) + ' All output')
            if not os.path.exists(newpath):
                os.makedirs(newpath)
            for i in LB.all_volunteers:
                if j.get() == 1:
                    LB.output_volunteer_detail(
                        i, os.path.join(newpath, i.get_name() + ' profile.xlsx', all=True))
                else:
                    LB.output_volunteer_detail(
                        i, os.path.join(newpath, i.get_name() + ' profile.xlsx'))
        else:
            if j.get() == 1:
                LB.output_volunteer_detail(
                    name, os.path.join('Output', str(datetime.today().date()) + ' ' + str(name) + ' profile.xlsx', all=True))
            else:
                LB.output_volunteer_detail(
                    name, os.path.join('Output', str(datetime.today().date()) + ' ' + str(name) + ' profile.xlsx'))
        status.config(text=f'Successfully output {name} profile.')
        top.destroy()
    end = Button(top, text='Select', command=end)
    end.pack(side='bottom')


button_10 = Button(
    option_frame, text='Output volunteer details', command=output_volunteer_det, width=20, height=2)
button_10.grid(column=3, row=3, ipady=5, pady=10, padx=5)


def suspend():
    global LB
    lst = list(map(lambda x: x.get_name(), filter(
        lambda x: x.suspended == False, LB.all_elderly)))
    top = Toplevel()
    top.resizable(False, False)
    top.title('Select person')
    top.geometry('250x175')
    option = StringVar(top)
    option.set('Choose person')
    choose = OptionMenu(top, option, *lst)
    choose.place(relx=0.5, rely=0.5, anchor='center')

    def end():
        name = option.get()
        if name == 'Choose person':
            messagebox.showerror('LB planner', 'No person selected.')
            return
        LB.suspend(name)
        status.config(text=f'{name} successfully suspended.')
        top.destroy()
    end = Button(top, text='Select', command=end)
    end.pack(side='bottom')


button_11 = Button(option_frame, text='Suspend elderly',
                   command=suspend, width=20, height=2)
button_11.grid(column=1, row=4, ipady=5, pady=10, padx=5)


def unsuspend():
    global LB
    lst = list(map(lambda x: x.get_name(), filter(
        lambda x: x.suspended == True, LB.all_elderly)))
    if not lst:
        messagebox.showerror('LB planner', 'No person to unsuspend.')
        return
    top = Toplevel()
    top.resizable(False, False)
    top.title('Select person')
    top.geometry('250x175')
    option = StringVar(top)
    option.set('Choose person')
    choose = OptionMenu(top, option, *lst)
    choose.place(relx=0.5, rely=0.5, anchor='center')

    def end():
        name = option.get()
        if name == 'Choose person':
            messagebox.showerror('LB planner', 'No person selected.')
            return
        LB.unsuspend(name)
        status.config(text=f'{name} successfully unsuspended.')
        top.destroy()
    end = Button(top, text='Select', command=end)
    end.pack(side='bottom')


button_12 = Button(
    option_frame, text='Unsuspend elderly', command=unsuspend, width=20, height=2)
button_12.grid(column=2, row=4, ipady=5, pady=10, padx=5)


def upd_avail():
    global LB
    top = Toplevel()
    top.title('Choose file')
    top.geometry('250x175')
    top.resizable(False, False)

    def open():
        filename = filedialog.askopenfilename(
            initialdir='.', title='Select file', filetypes=[("Microsoft Excel Worksheet", "*.xlsx")])
        try:
            LB.update_availability(filename)
        except ValueError as e:
            messagebox.showerror('LB planner', e)
            return
        status.config(text=f"Volunteers' availability successfully updated.")
        top.destroy()
    B1 = Button(top, text='Select file', command=open)
    B1.place(relx=0.5, rely=0.5, anchor='center')


button_13 = Button(option_frame, text='Update availability',
                   command=upd_avail, width=20, height=2)
button_13.grid(column=1, row=2, ipady=5, pady=10, padx=5)


def output_att():
    global LB
    response = messagebox.askyesno(
        'LB planner', f'Output the attendances of all volunteers?')
    if response:
        LB.output_attendances(os.path.join('Output', str(
            datetime.today().date()) + ' attendances.xlsx'))
        status.config(text=f'Successfully output attendances.')


button_14 = Button(
    option_frame, text='Output attendances', command=output_att, width=20, height=2)
button_14.grid(column=4, row=3, ipady=5, pady=10, padx=5)


def motivate():
    # max 120 chars long, optimally 80. One line = 40 chars.
    quotes = ['\(-ㅂ-)/ ♥ ♥ ♥', '♡＾▽＾♡', '（*＾3＾）', '(๑╹ڡ╹)╭ ～ ♡', '(=^ェ^=) ', 'ʕ •ᴥ•ʔ', '(づ｡◕‿‿◕｡)づ', '｡◕ ‿ ◕｡ ',
              'ヾ(⌐■_■)ノ♪', 'ᕙ( ͡° ͜ʖ ͡°)ᕗ', '＼(＾O＾)／', '♪┏(・o･)┛♪┗ ( ･o･) ┓♪', '└[∵┌]└[ ∵ ]┘[┐∵]┘', '(ﾉ^_^)ﾉ', '(ノ・∀・)ノ', '(✿◕‿◕✿)']
    i = randint(0, len(quotes)-1)
    status.config(text=quotes[i])
    return


button_15 = Button(option_frame, text=':D',
                   command=motivate, width=20, height=2)
button_15.grid(column=3, row=4, ipady=5, pady=10, padx=5)


def end_year():
    global LB
    response1 = messagebox.askyesno(
        'LB planner', f'Confirm the end of current year?')
    if response1:
        response2 = messagebox.askokcancel(
            'LB planner', f'Are you REALLY sure you want to end the current year?')
    if response1 and response2:
        LB.end_year()
        status.config(
            text='Year successfully ended. Hope that you had meaningful fun in your PD journey :)')


button_16 = Button(option_frame, text='End year',
                   command=end_year, width=20, height=2)
button_16.grid(column=4, row=4, ipady=5, pady=10, padx=5)


def exit():
    save()
    root.destroy()


exit_button = Button(root, text='Exit', font='Helvetica, 10',
                     fg='red', command=exit, width=3)
exit_button.pack(side='bottom')

if not isinstance(SUNDAY, Session):
    if SUNDAY != None:
        messagebox.showerror(
            'LB planner', 'Invalid session data found in Data/session_data.pkl')
        root.destroy()

if not isinstance(LB, Planner):
    messagebox.showerror(
        'LB planner', 'Invalid planner data found in Data/planner_data.pkl')
    root.destroy()

paths = [os.path.join('Archive', 'Previous Detailing'), 'Output']
for i in paths:
    if not os.path.exists(i):
        os.makedirs(i)

root.mainloop()
