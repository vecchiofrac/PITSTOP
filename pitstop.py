# -*- coding: utf-8 -*-
# 
# verifica le presenze memorizzate in Pitstop
# implementazione con pySimpleGUI

# Franzsoft 2022 (:)

import os
import re
import PySimpleGUI as sg
from datetime import datetime, timedelta
from sys import exit
from Dbpitstop import Database
import operator

__version__ = "2.4 (2022 - pySimpleGUI)"

# -------------------------------------------------------------------------------

def validate(text):
    regex1 = "\d{1,4}/\d{1,2}/\d{2}"        # accept two-digit year
    regex2 = "\d{1,4}/\d{1,2}/\d{4}"        # accept four-digit year
    result = re.match(regex1, text) or re.match(regex2, text)
    #return False if result is None or result.group() != text else True
    return False if result is None else True


def get_recordset(date):    
    sql = f"""
            SELECT Cognome as user, Format(ApptStart, 'HH\:MM'), Format(ApptEnd, 'HH\:MM'), ApptSubject, ApptLocation, ApptNotes, ApptBookings
            FROM tblAppointments a 
            INNER JOIN tbUtenti b ON a.idutente = b.IDUTENTE
            WHERE DateValue(?) Between DateValue(ApptStart) And DateValue(ApptEnd) 
            ORDER BY Cognome, DateValue(ApptStart), TimeValue(ApptStart); 
            """

    rs = db.get_rs(sql, (date,))
    return [] if rs is None else rs


def get_presenze_at(date):
    rs = get_recordset(date)
    if rs is []:
        return []
    else:
        m=[]
        for v in rs:
            l=[ v[0], v[1], v[2], v[3], v[4], v[5], get_bookings(v[6]) ]
            m.append(l)
        
        return m

def get_bookings(tmp):
    prenots = ['Auto', 'PC1', 'Pro1', 'Sala r.', 'Dinam.', 'PC2', 'Pro2', 'Telec.']
    zipped = zip(prenots, tmp)
    s=[]
    try:
        [s.append(zi[0]) for zi in list(zipped) if int(zi[1])]
    except Exception as e:
        print(e)
        return ''

    return '+'.join(s)




def sort_table(table, cols):
    """ sort a table by multiple columns
        table: a list of lists (or tuple of tuples) where each inner list
               represents a row
        cols:  a list (or tuple) specifying the column numbers to sort by
               e.g. (1,0) would sort by column 1, then by column 0
    """
    for col in reversed(cols):
        try:
            table = sorted(table, key=operator.itemgetter(col))
        except Exception as e:
            sg.popup_error('Error in sort_table', 'Exception in sort_table', e)
    return table


# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
def main():

    # --------------------------------------------------------
    def update(d):
        window['-VIEW-'].update(d)
        window.set_title(f'{title} del {d}')
        window['-TABLE-'].update(values=get_presenze_at(d))
    # --------------------------------------------------------

    sg.theme('GreenMono')
    title = "Pitstop registro presenze"
    
    d = datetime.strftime(datetime.today(), "%d/%m/%Y")    # data odierna in formato testo

    headings = ['User', 'Dalle', 'Alle', 'Motivazione', 'Località', 'Descrizione', 'Prenotazioni']
    col_widths = [13, 7, 7, 15, 15, 34, 10]

    table_values = get_presenze_at(d)

    layout = [
                [ sg.Table(
                    values=table_values, 
                    key='-TABLE-', 
                    headings=headings, 
                    font="Consolas 10",
                    auto_size_columns=False,
                    hide_vertical_scroll=True,
                    def_col_width=20,
                    num_rows=14, 
                    col_widths=col_widths,
                    background_color='white', 
                    alternating_row_color='lightyellow', 
                    vertical_scroll_only=False, 
                    cols_justification=('l','c','c','l', 'l', 'l', 'r'),
                    selected_row_colors='red on yellow',
                    expand_x=True,
                    expand_y=False,
                    bind_return_key=True,
                    #enable_events=True,
                    select_mode=sg.TABLE_SELECT_MODE_BROWSE
                    ),
                ],
                [ sg.Text('Showing date:'), sg.Text(d, key='-VIEW-'), 
                  sg.Text('Enter new date:'), sg.InputText(key='-DATE-', size=(10,1), enable_events=True),
                  sg.Button('Giorno prec.', key='-PREV-'), 
                  sg.Button(' -  OGGI  - ', key='-TODAY-', button_color="black on yellow", font="Consolas 12 bold"), 
                  sg.Button('Giorno succ.', key='-NEXT-'), 
                  sg.InputText(key='-CAL-', enable_events=True, visible=False),
                  sg.CalendarButton('Calendario', target='-CAL-', format='%d/%m/%Y', key='-CALENDAR-')
                ]
             ] 

    window = sg.Window(title, layout, return_keyboard_events=True, use_default_focus=True, finalize=True, size=(1000,300), font="Consolas 12")
    window.move(500, 300)
    window['-DATE-'].bind("<Return>", "_Enter")

    # -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -
    while True:
        event, values = window.read(timeout=100)
        
        if event in [sg.WIN_CLOSED, "Escape:27"]:  # this spares me a 'Quit' button!
            break

        elif event == '-DATE-' and len(values['-DATE-']) and values['-DATE-'][-1] not in ('0123456789/'):
            # delete last char from input if invalid
            window['-DATE-'].update(values['-DATE-'][:-1])

        elif event == '-DATE-_Enter':
            text = values['-DATE-']
            if validate(text):
                d, m, y = text.split('/')
                d = f'{d:0>2}'
                m = f'{m:0>2}'
                y = int(y)
                y = f'{y+2000:0>4}' if y<100 else f'{y:0>4}'
                text = "/".join([d, m, y])
                window['-VIEW-'].update(text)
                window['-DATE-'].update('')
                window['-TABLE-'].update(values=get_presenze_at(text))

            else:
                window['-DATE-'].Widget.select_range(0, 'end')
                window['-DATE-'].Widget.icursor('end')
                sg.popup_error('Invalid date')

        if event == '-TABLE-':
                index = values[event][0]
                date = window['-VIEW-'].get()
                data = get_presenze_at(date)
                loc = data[index][4] 
                sg.popup(f"{data[index][0]} >> {data[index][3]} {'' if loc=='' else 'a '+loc}\n{data[index][5]}")

            
        elif event == '-PREV-':
            d = str(datetime.strftime(datetime.strptime(window['-VIEW-'].get(), '%d/%m/%Y') + timedelta(days=-1), '%d/%m/%Y'))
            update(d)

        elif event == '-TODAY-':
            d = str(datetime.strftime(datetime.today(), '%d/%m/%Y'))
            update(d)

        elif event == '-NEXT-':
            d = str(datetime.strftime(datetime.strptime(window['-VIEW-'].get(), '%d/%m/%Y') + timedelta(days=1), '%d/%m/%Y'))
            update(d)

        elif event == '-CAL-':
            d = values['-CAL-']
            update(d)

    # -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -   -            

    window.close(); del window
    print("Bye.")



# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
if __name__ == '__main__':
    db = Database()
    main()
    db.terminate()
