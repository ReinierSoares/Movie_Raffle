import PySimpleGUI as sg
import pandas as pd
from random import randint
from googleapiclient.discovery import build
import webbrowser
import openpyxl


def random_film(genero):
    global window
    global link

    while True:
        random_number = randint(0, len(movie_list)+1)
        if genero == 'Todos':
            break
        else:
            if genero == gender_list[random_number]:
                break
    film = movie_list[random_number]
    link = link_list[random_number]
    window["-FILM-"].update(film)


def add_film_window():
    global window
    items = ['Romance', 'Terror', 'Heróis', 'Animação',
             'Ação/Aventura', 'Suspense']
    layout = [[sg.Menu(menu_layout)],
              [sg.Text('Selecione um gênero:')],
              [sg.Combo(values=items, key='-GENDER-', size=(20, 1))],
              [sg.Text('Nome do Filme:'), sg.InputText(key='-FILM-')],
              [sg.Text('Link do Filme:'), sg.InputText(key='-LINK-')],
              [sg.Button('Ok'), sg.Button('Cancelar')]]

    window = sg.Window('Adicionar Filme', layout)

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Sair':
            break
        elif event == 'Cancelar':
            window.close()
            main_window()
        elif event == 'Ok':
            gender_value = values['-GENDER-']
            film_value = values['-FILM-']
            link_value = values['-LINK-']

            if film_value == '':
                pass
            else:
                add_movie(gender_value, film_value, link_value)
        menu_bar(event)


def open_film_window(movie, link, check, row):
    global table

    if check.lower() == '✔️':
        check_button = sg.Button('Desmarcar visto')
    else:
        check_button = sg.Button('Marcar visto')

    layout = [[sg.Text(movie)],
              [sg.Button('Abrir Link'), check_button, sg.Button('Cancelar')]]

    window = sg.Window(f'Filme').Layout(layout)

    # Mostra a janela e aguarda a resposta do usuário
    button, _ = window.Read()

    # Verifica qual botão foi pressionado
    if button == 'Abrir Link':
        try:
            webbrowser.open(link)
        except:
            sg.popup(
                f"O filme {movie} não possui um link!", title=movie)

    elif button == 'Marcar visto':
        workbook = openpyxl.load_workbook('film_list.xlsx')

        worksheet = workbook.active

        worksheet[f'D{row+2}'] = "Sim"

        workbook.save('film_list.xlsx')
        att_movies()
    elif button == 'Desmarcar visto':
        workbook = openpyxl.load_workbook('film_list.xlsx')

        worksheet = workbook.active

        worksheet[f'D{row+2}'] = "Não"

        workbook.save('film_list.xlsx')
        att_movies()
    else:
        pass
    df = pd.read_excel("film_list.xlsx")
    column_to_remove = df['Link']
    df = df.drop(columns=['Link'])

    data_list = df.values.tolist()
    window.Close()
    return data_list


def data_check_list(data_list):
    for linha in data_list:
        if linha[2].lower() == 'sim':  # verifique o valor da segunda coluna
            linha.insert(2, '✔️')
        else:
            linha.insert(2, '☐')
    return data_list


def list_window():
    global window

    df = pd.read_excel("film_list.xlsx")
    column_to_remove = df['Link']
    df = df.drop(columns=['Link'])

    data_list = df.values.tolist()
    data = data_check_list(data_list)

    layout = [
        [sg.Menu(menu_layout)],
        [sg.InputText(key='-SEARCH-'), sg.Button('Pesquisar'),
         sg.Button('Limpar')],
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

    window = sg.Window("Tabela de Filmes", auto_size_text=True, auto_size_buttons=True,
                       grab_anywhere=False, resizable=True,
                       layout=layout, finalize=True)
    window['-TABLE-'].expand(True, True)
    window['-TABLE-'].table_frame.pack(expand=True, fill='both')
    while True:
        event, values = window.read()

        if event == 'Sair' or event == sg.WIN_CLOSED:
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
        elif event == 'Pesquisar':

            search_term = values['-SEARCH-']
            if search_term:
                filtered_data = df[df.apply(lambda x: search_term.lower(
                ) in x.astype(str).str.lower().values.tolist(), axis=1)]
                data_list = filtered_data.values.tolist()

                window['-TABLE-'].update(values=data_list)
            else:
                data_list = df.values.tolist()

                window['-TABLE-'].update(values=data_list)
        elif event == 'Limpar':
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
              [sg.Text('Selecione um gênero:')],
              [sg.Combo(values=items, key='-GENDER-', size=(20, 1)),
               sg.Button('Sortear')],

              [sg.Text("Filme:"), sg.Output(size=(50, 1), key="-FILM-")],
              [sg.Button('Open link')]]

    # Create the window
    window = sg.Window('Sorteador de Filmes', layout)

    while True:
        event, values = window.read()
        selected_item = values['-GENDER-']

        if event == 'Sortear':
            if selected_item != '':
                random_film(selected_item)
        elif event == 'Open link':
            if link is None:
                pass
            else:

                webbrowser.open(link)
        elif event == sg.WIN_CLOSED or event == 'Sair':
            break
        menu_bar(event)

    # Close the window
    window.close()


def add_movie(gender, film, link):

    for movie in movie_list:
        if movie.lower() == film.lower():
            sg.popup(f'O filme ({film}) ja está na lista!')
            return

    row = len(movie_list) + 2
    # Load the existing workbook
    workbook = openpyxl.load_workbook('film_list.xlsx')

    # Select the worksheet you want to add data to
    worksheet = workbook.active

    # Add data to the worksheet
    worksheet[f'A{row}'] = gender
    worksheet[f'B{row}'] = film.replace('\n', '').title()
    worksheet[f'C{row}'] = link
    worksheet[f'D{row}'] = 'Não'

    # Save the changes to the workbook
    try:
        workbook.save('film_list.xlsx')
        att_movies()

        sg.popup(f'O filme ({film.title()}) foi adicionado!')
    except (PermissionError):
        sg.popup('É necessário fechar o arquivo Excel!')
    except Exception as e:
        sg.popup(e)


def edit_movie_window():
    global window
    movies = []
    for movie in movie_list:
        movies.append(movie)
    genders = ['Romance', 'Terror', 'Heróis', 'Animação',
               'Ação/Aventura', 'Suspense']
    layout = [[sg.Menu(menu_layout)],
              [sg.Text('Filme que deseja editar:')],
              [sg.Combo(values=movies, key='-COMBO-', size=(50, 1))],
              [sg.Text('Editar o Gênero do Filme:')],
              [sg.Combo(values=genders, key='-GENDER-', size=(20, 1))],
              [sg.Text('Editar Nome do Filme:'), sg.InputText(key='-FILM-')],
              [sg.Text('Editar Link do Filme:'), sg.InputText(key='-LINK-')],
              [sg.Button('Ok'), sg.Button('Cancelar')]]

    window = sg.Window('Editar Filme', layout)

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Sair':
            break
        elif event == 'Cancelar':
            window.close()
            main_window()
        elif event == 'Ok':
            selected_movie = values['-COMBO-']
            gender_value = values['-GENDER-']
            film_value = values['-FILM-']
            link_value = values['-LINK-']

            if selected_movie != '':
                if film_value != '' or link_value != '' or gender_value != '':
                    edit_movie(selected_movie, gender_value,
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
    genders = ['Romance', 'Terror', 'Heróis', 'Animação',
               'Ação/Aventura', 'Suspense']
    layout = [[sg.Menu(menu_layout)],
              [sg.Text('Filme que deseja Deletar:')],
              [sg.Combo(values=movies, key='-COMBO-', size=(50, 1))],
              [sg.Button('Ok'), sg.Button('Cancelar')]]

    window = sg.Window('Deletar Filme', layout)

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Sair':
            break
        elif event == 'Cancelar':
            window.close()
            main_window()
        elif event == 'Ok':
            selected_movie = values['-COMBO-']
            delete_movie(selected_movie)
            window['-COMBO-'].update()

        menu_bar(event)
    window.close()


def edit_movie(selected_movie, gender, film, link):
    workbook = openpyxl.load_workbook('film_list.xlsx')

    # Select the worksheet you want to add data to
    worksheet = workbook.active

    # Add data to the worksheet

    for index, movie in enumerate(movie_list):
        if movie.lower() == selected_movie.lower():
            index += 2
            if gender != '':
                worksheet[f'A{index}'] = gender
            if film != '':
                worksheet[f'B{index}'] = film.replace('\n', '').title()
            if link != '':
                worksheet[f'C{index}'] = link

    # Save the changes to the workbook
    try:
        workbook.save('film_list.xlsx')
        att_movies()
        sg.popup(f'O filme ({selected_movie.title()}) foi Editado!')
    except (PermissionError):
        sg.popup('É necessário fechar o arquivo Excel!')
    except Exception as e:
        sg.popup(e)


def delete_movie(selected_movie):
    workbook = openpyxl.load_workbook('film_list.xlsx')

    for index, movie in enumerate(movie_list):
        if movie.lower() == selected_movie.lower():
            worksheet = workbook['List']
            # Excluir a linha
            worksheet.delete_rows(index+2)

    try:
        workbook.save('film_list.xlsx')
        att_movies()
        sg.popup(f'O filme ({selected_movie.title()}) foi Deletado!')
    except (PermissionError):
        sg.popup('É necessário fechar o arquivo Excel!')
    except Exception as e:
        sg.popup(e)


def att_movies():
    global table
    global movie_list
    global gender_list
    global link_list

    try:
        df = pd.read_excel("film_list.xlsx", sheet_name='List')
        df_sorted = df.sort_values('Filme')
        df_sorted.to_excel('film_list.xlsx', index=False, sheet_name='List')

        table = pd.read_excel("film_list.xlsx", None)

        movie_list = table['List']['Filme']
        gender_list = table['List']['Gênero']
        link_list = table['List']['Link']
    except (PermissionError):
        sg.popup('É necessário fechar o arquivo Excel!')
    except Exception as e:
        sg.popup(e)


def menu_bar(event):
    if event == 'Adicionar':
        window.close()
        add_film_window()
    elif event == 'Sorteador':
        window.close()
        main_window()
    elif event == 'Lista':
        window.close()
        list_window()
    elif event == 'Editar':
        window.close()
        edit_movie_window()
    elif event == 'Excluir':
        window.close()
        delete_movie_window()
    elif event == 'About':
        sg.popup('Produced by: Reinier Soares')


table = pd.read_excel("film_list.xlsx", None)

menu_layout = [
    ['Filme', ['Sorteador', 'Lista', 'Adicionar', 'Editar', 'Excluir', 'Sair']], ['Help', ['About']]]

window = None
link = None
movie_list = table['List']['Filme']
gender_list = table['List']['Gênero']
link_list = table['List']['Link']

main_window()
