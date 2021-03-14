import tkinter
import tkcalendar
from datetime import datetime



def add_lables(event):
    date  = C.selection_get()
    temp.set(datetime.strftime(date,'%d %B %Y'))
    lable.set('')
    tkinter.Label(Left,textvariable=temp).grid(row=0)
    try:
        eventid = C.get_calevents(date)
        for i in eventid:
            lable.set(lable.get()+C.calevent_cget(i,'text')+'\n')
        
    except:
        lable.set('')
    tkinter.Label(Left,textvariable=lable).grid(row=1)



top = tkinter.Tk()
colour_scheme = [()]
C = tkcalendar.Calendar(top,background='darkviolet',selectbackground='lightslateblue',normalbackground  ='slateblue',weekendbackground='steelblue1' ,headersbackground ='lightblue1',othermonthbackground ='slateblue4',othermonthwebackground ='slateblue4')
Left = tkinter.Frame(top,highlightbackground="black",highlightthickness=1,height = 185,width =150)
temp = tkinter.StringVar()
lable = tkinter.StringVar()

schedule = open('C:/Users\Chua Wei Yang\Desktop\Project\Ouroborus\Data\Scheduel\schedule.dat','rb')
for i in schedule:
    i = i.decode()
    lst = i.split(' ')
    text = lst[2]
    date =lst[1].split('-')
    
    C.calevent_create(datetime(int(date[0]),int(date[1]),int(date[2])),text,[1])


C.grid(column=0,row=0)
Left.grid(column=1,row=0)
Left.grid_propagate(0)

top.bind('<Button 1>',add_lables)
top.mainloop()