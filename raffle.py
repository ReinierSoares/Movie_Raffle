import PySimpleGUI as sg
import pandas as pd
from random import randint
import webbrowser
import openpyxl


def random_film(genre):
    global window
    global link

    while True:
        random_number = randint(0, len(movie_list)+1)
        if genre == 'Todos':
            break
        else:
            if genre == genre_list[random_number]:
                break
    film = movie_list[random_number]
    link = link_list[random_number]
    window["-FILM-"].update(film)


def add_film_window():
    global window
    items = ['Romance', 'Terror', 'Heróis', 'Animação',
             'Ação/Aventura', 'Suspense']
    layout = [[sg.Menu(menu_layout)],
              [sg.Text('Select a genre:')],
              [sg.Combo(values=items, key='-GENRE-', size=(20, 1))],
              [sg.Text('Movie Name:'), sg.InputText(key='-FILM-')],
              [sg.Text('Movie Link:'), sg.InputText(key='-LINK-')],
              [sg.Button('Ok'), sg.Button('cancel')]]

    window = sg.Window('Add Movie', layout)

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Exit':
            break
        elif event == 'cancel':
            window.close()
            main_window()
        elif event == 'Ok':
            genre_value = values['-GENRE-']
            film_value = values['-FILM-']
            link_value = values['-LINK-']

            if film_value == '':
                pass
            else:
                add_movie(genre_value, film_value, link_value)
        menu_bar(event)


def open_film_window(movie, link, check, row):
    global table

    if check.lower() == '✔️':
        check_button = sg.Button('Uncheck')
    else:
        check_button = sg.Button('Check')

    layout = [[sg.Text(movie)],
              [sg.Button('Open Link'), check_button, sg.Button('cancel')]]

    window = sg.Window(f'Movie').Layout(layout)

    # Mostra a janela e aguarda a resposta do usuário
    button, _ = window.Read()

    # Verifica qual botão foi pressionado
    if button == 'Open Link':
        try:
            webbrowser.open(link)
        except:
            sg.popup(
                f"The movie {movie} doesn't have a link.", title=movie)

    elif button == 'Check':
        workbook = openpyxl.load_workbook('movie_list.xlsx')

        worksheet = workbook.active

        worksheet[f'D{row+2}'] = "Sim"

        workbook.save('movie_list.xlsx')
        att_movies()
    elif button == 'Uncheck':
        workbook = openpyxl.load_workbook('movie_list.xlsx')

        worksheet = workbook.active

        worksheet[f'D{row+2}'] = "Não"

        workbook.save('movie_list.xlsx')
        att_movies()
    else:
        pass
    df = pd.read_excel("movie_list.xlsx")
    column_to_remove = df['Link']
    df = df.drop(columns=['Link'])

    data_list = df.values.tolist()
    window.Close()
    return data_list


def data_check_list(data_list):
    for row in data_list:
        if row[2].lower() == 'sim':  # verifique o valor da segunda coluna
            row.insert(2, '✔️')
        else:
            row.insert(2, '☐')
    return data_list


def list_window():
    global window

    df = pd.read_excel("movie_list.xlsx")
    column_to_remove = df['Link']
    df = df.drop(columns=['Link'])

    data_list = df.values.tolist()
    data = data_check_list(data_list)

    layout = [
        [sg.Menu(menu_layout)],
        [sg.InputText(key='-SEARCH-'), sg.Button('Search'),
         sg.Button('Clean')],
        [sg.Table(
            values=data,
            headings=df.columns.tolist(),
            max_col_width=70,
            auto_size_columns=True,
            justification='left',
            num_rows=min(len(data_list), 20),
            row_height=30,
            key='-TABLE-',
            enable_events=True
        )]
    ]

    window = sg.Window("Movie Table", auto_size_text=True, auto_size_buttons=True,
                       grab_anywhere=False, resizable=True,
                       layout=layout, finalize=True)
    window['-TABLE-'].expand(True, True)
    window['-TABLE-'].table_frame.pack(expand=True, fill='both')
    while True:
        event, values = window.read()

        if event == 'Exit' or event == sg.WIN_CLOSED:
            break
        if event == '-TABLE-':

            if values['-TABLE-']:
                selected_row = values['-TABLE-'][0]
                table_movie = data_list[selected_row][1]
                table_link = column_to_remove[selected_row]
                table_check = data_list[selected_row][2]
                data_list = open_film_window(
                    table_movie, table_link, table_check, selected_row)
                data = data_check_list(data_list)
                window['-TABLE-'].update(values=data)
        elif event == 'Search':

            search_term = values['-SEARCH-']
            if search_term:
                filtered_data = df[df.apply(lambda x: search_term.lower(
                ) in x.astype(str).str.lower().values.tolist(), axis=1)]
                data_list = filtered_data.values.tolist()

                window['-TABLE-'].update(values=data_list)
            else:
                data_list = df.values.tolist()

                window['-TABLE-'].update(values=data_list)
        elif event == 'Clean':
            data_list = df.values.tolist()

            window['-TABLE-'].update(values=data_list)
        menu_bar(event)

    window.close()


def main_window():
    global window
    global link
    # Define the dictionary of items
    items = ['Todos', 'Romance', 'Terror', 'Heróis', 'Animação',
             'Ação/Aventura', 'Suspense']
    link = None

    # Define the layout of your window
    layout = [[sg.Menu(menu_layout)],
              [sg.Text('Select a Movie Genre:')],
              [sg.Combo(values=items, key='-GENRE-', size=(20, 1)),
               sg.Button('raffle')],

              [sg.Text("Movie:"), sg.Output(size=(50, 1), key="-FILM-")],
              [sg.Button('Open link')]]

    # Create the window
    window = sg.Window('Movie_Raffle', layout)

    while True:
        event, values = window.read()
        selected_item = values['-GENRE-']
        if event == sg.WIN_CLOSED or event == 'Exit':
            break
        elif event == 'raffle':
            if selected_item != '':
                random_film(selected_item)
        elif event == 'Open link':
            if link is None:
                pass
            else:
                webbrowser.open(link)

        menu_bar(event)
    # Close the window
    window.close()


def add_movie(genre, film, link):

    for movie in movie_list:
        if movie.lower() == film.lower():
            sg.popup(f'The movie ({film}) is already on the list.')
            return

    row = len(movie_list) + 2
    # Load the existing workbook
    workbook = openpyxl.load_workbook('movie_list.xlsx')

    # Select the worksheet you want to add data to
    worksheet = workbook.active

    # Add data to the worksheet
    worksheet[f'A{row}'] = genre
    worksheet[f'B{row}'] = film.replace('\n', '').title()
    worksheet[f'C{row}'] = link
    worksheet[f'D{row}'] = 'Não'

    # Save the changes to the workbook
    try:
        workbook.save('movie_list.xlsx')
        att_movies()

        sg.popup(f'The movie ({film.title()}) was added.')
    except (PermissionError):
        sg.popup('It is necessary to close the Excel file..')
    except Exception as e:
        sg.popup(e)


def edit_movie_window():
    global window
    movies = []
    for movie in movie_list:
        movies.append(movie)
    genres = ['Romance', 'Terror', 'Heróis', 'Animação',
              'Ação/Aventura', 'Suspense']
    layout = [[sg.Menu(menu_layout)],
              [sg.Text('Movie you want to edit:')],
              [sg.Combo(values=movies, key='-COMBO-', size=(50, 1))],
              [sg.Text('Edit Movie Genre:')],
              [sg.Combo(values=genres, key='-GENRE-', size=(20, 1))],
              [sg.Text('Edit Movie Name:'), sg.InputText(key='-FILM-')],
              [sg.Text('Edit Movie Link:'), sg.InputText(key='-LINK-')],
              [sg.Button('Ok'), sg.Button('cancel')]]

    window = sg.Window('Edit Movie', layout)

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Exit':
            break
        elif event == 'cancel':
            window.close()
            main_window()
        elif event == 'Ok':
            selected_movie = values['-COMBO-']
            genre_value = values['-GENRE-']
            film_value = values['-FILM-']
            link_value = values['-LINK-']

            if selected_movie != '':
                if film_value != '' or link_value != '' or genre_value != '':
                    edit_movie(selected_movie, genre_value,
                               film_value, link_value)
                    window['-COMBO-'].update()

        menu_bar(event)
    window.close()


def delete_movie_window():
    global window
    global movie_list
    movies = []
    for movie in movie_list:
        movies.append(movie)
    genres = ['Romance', 'Terror', 'Heróis', 'Animação',
              'Ação/Aventura', 'Suspense']
    layout = [[sg.Menu(menu_layout)],
              [sg.Text('Movie you want to Delete:')],
              [sg.Combo(values=movies, key='-COMBO-', size=(50, 1))],
              [sg.Button('Ok'), sg.Button('cancel')]]

    window = sg.Window('Delet Movie', layout)

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Exit':
            break
        elif event == 'cancel':
            window.close()
            main_window()
        elif event == 'Ok':
            selected_movie = values['-COMBO-']
            delete_movie(selected_movie)
            window['-COMBO-'].update()

        menu_bar(event)
    window.close()


def edit_movie(selected_movie, genre, film, link):
    workbook = openpyxl.load_workbook('movie_list.xlsx')

    # Select the worksheet you want to add data to
    worksheet = workbook.active

    # Add data to the worksheet

    for index, movie in enumerate(movie_list):
        if movie.lower() == selected_movie.lower():
            index += 2
            if genre != '':
                worksheet[f'A{index}'] = genre
            if film != '':
                worksheet[f'B{index}'] = film.replace('\n', '').title()
            if link != '':
                worksheet[f'C{index}'] = link

    # Save the changes to the workbook
    try:
        workbook.save('movie_list.xlsx')
        att_movies()
        sg.popup(f'The movie({selected_movie.title()}) was edited.')
    except (PermissionError):
        sg.popup('It is necessary to close the Excel file.')
    except Exception as e:
        sg.popup(e)


def delete_movie(selected_movie):
    workbook = openpyxl.load_workbook('movie_list.xlsx')

    for index, movie in enumerate(movie_list):
        if movie.lower() == selected_movie.lower():
            worksheet = workbook['List']
            # Delete a linha
            worksheet.delete_rows(index+2)

    try:
        workbook.save('movie_list.xlsx')
        att_movies()
        sg.popup(f'The movie ({selected_movie.title()}) has been Deleted.')
    except (PermissionError):
        sg.popup('It is necessary to close the Excel file.')
    except Exception as e:
        sg.popup(e)


def att_movies():
    global table
    global movie_list
    global genre_list
    global link_list

    try:
        df = pd.read_excel("movie_list.xlsx", sheet_name='List')
        df_sorted = df.sort_values('Filme')
        df_sorted.to_excel('movie_list.xlsx', index=False, sheet_name='List')

        table = pd.read_excel("movie_list.xlsx", None)

        movie_list = table['List']['Filme']
        genre_list = table['List']['Gênero']
        link_list = table['List']['Link']
    except (PermissionError):
        sg.popup('It is necessary to close the Excel file.')
    except Exception as e:
        sg.popup(e)


def menu_bar(event):
    if event == 'Add':
        window.close()
        add_film_window()
    elif event == 'Raffle':
        window.close()
        main_window()
    elif event == 'List':
        window.close()
        list_window()
    elif event == 'Edit':
        window.close()
        edit_movie_window()
    elif event == 'Delete':
        window.close()
        delete_movie_window()
    elif event == 'About':
        sg.popup('Produced by: Reinier Soares')


table = pd.read_excel("movie_list.xlsx", None)

menu_layout = [
    ['Movie', ['Raffle', 'Add', 'List',  'Edit', 'Delete', 'Exit']], ['Help', ['About']]]

window = None
link = None
movie_list = table['List']['Filme']
genre_list = table['List']['Gênero']
link_list = table['List']['Link']

main_window()
