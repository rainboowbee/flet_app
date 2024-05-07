import xml.etree.ElementTree as ET
import pandas as pd

def parse_xml_to_dataframe(xml_file_path):
    # Загружаем и разбираем XML файл
    tree = ET.parse(xml_file_path)
    root = tree.getroot()

    players_id_data = []

    for player in root.findall('./Tournament/Players/Player'):
        player_id = player.get('id')
        players_id_data.append(player_id)
    
    print(players_id_data)


    # Парсим информацию о игроках
    players_data = []
    for idx, player in enumerate(root.findall('.//Player')):
        player_unique_id = player.attrib['id']
        name = player.find('name').text if player.find('name') is not None else None
        location = player.find('location').text if player.find('location') is not None else None
        birthdate = player.find('birthdate').text if player.find('birthdate') is not None else None
        rating = player.find('rating').text if player.find('rating') is not None else None

        print(player_unique_id, name, location, birthdate, rating)

        # Проверяем, есть ли необходимая информация перед добавлением в список
        if name and location and birthdate and rating and player_unique_id in players_id_data:
            player_info = {
                '№': idx + 1,  # Индексация начинается с 1
                'ФИО': name,
                'Город': location,
                'Год рождения': birthdate,
                'рейтинг': rating
            }
            players_data.append(player_info)
    players_df = pd.DataFrame(players_data)
    
    # Парсим информацию о турнире
    tournament_data = []
    tournament = root.find('.//Tournament')
    if tournament is not None:
        header = tournament.find('Header')
        if header is not None:
            tournament_info = {
                'Дата': header.find('date').text if header.find('date') is not None else '',
                'Название': header.find('name').text if header.find('name') is not None else '',
                'Адрес': header.find('addr').text if header.find('addr') is not None else '',
                'Организатор': header.find('organizer').text if header.find('organizer') is not None else '',
                'Количество игроков': header.find('numPlayers').text if header.find('numPlayers') is not None else '',
                'Колличество столов': header.find('numTables').text if header.find('numTables') is not None else ''
            }
            tournament_data.append(tournament_info)

    # Преобразуем данные о турнире в формат DataFrame
    tournament_df = pd.DataFrame(tournament_data)
    
    return players_df, tournament_df

def parse_xml_and_save_to_excel(xml_file_path, excel_file_path):
    # Загружаем и разбираем XML файл
    tree = ET.parse(xml_file_path)
    root = tree.getroot()

    players_id_data = []

    for player in root.findall('./Tournament/Players/Player'):
        player_id = player.get('id')
        players_id_data.append(player_id)
    
    # Парсим информацию о игроках
    players_data = []
    for idx, player in enumerate(root.findall('.//Player')):
        player_unique_id = player.attrib['id']
        name = player.find('name').text if player.find('name') is not None else None
        location = player.find('location').text if player.find('location') is not None else None
        birthdate = player.find('birthdate').text if player.find('birthdate') is not None else None
        rating = player.find('rating').text if player.find('rating') is not None else None

        # Проверяем, есть ли необходимая информация перед добавлением в список
        if name and location and birthdate and rating and player_unique_id in players_id_data:
            player_info = {
                '№': idx + 1,  # Индексация начинается с 1
                'ФИО': name,
                'Город': location,
                'Год рождения': birthdate,
                'рейтинг': rating
            }
            players_data.append(player_info)
    players_df = pd.DataFrame(players_data)
    
    # Парсим информацию о турнире
    tournament_data = []
    tournament = root.find('.//Tournament')
    if tournament is not None:
        header = tournament.find('Header')
        if header is not None:
            tournament_info = {
                'Дата': header.find('date').text if header.find('date') is not None else '',
                'Название': header.find('name').text if header.find('name') is not None else '',
                'Адрес': header.find('addr').text if header.find('addr') is not None else '',
                'Организатор': header.find('organizer').text if header.find('organizer') is not None else '',
                'Количество игроков': header.find('numPlayers').text if header.find('numPlayers') is not None else '',
                'Колличество столов': header.find('numTables').text if header.find('numTables') is not None else ''
            }
            tournament_data.append(tournament_info)

    # Преобразуем данные о турнире в формат DataFrame
    tournament_df = pd.DataFrame(tournament_data)
    
    # Сохраняем в Excel с заданной шириной столбцов и выравниванием
    with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
        players_df.to_excel(writer, sheet_name='Players', index=False)
        tournament_df_formatted = pd.DataFrame(tournament_df.stack().reset_index())
        tournament_df_formatted = tournament_df_formatted.iloc[:, 1:]  # Игнорируем столбец с индексом
        tournament_df_formatted.to_excel(writer, sheet_name='Tournament', index=False, header=False)

        workbook = writer.book
        center_format = workbook.add_format({'align': 'center'})

        # Получаем объект worksheet, чтобы установить настройки форматирования
        worksheet_players = writer.sheets['Players']
        worksheet_players.set_column('A:A', 4)  # number
        worksheet_players.set_column('B:B', 35)  # name, 7.5cm approx 28
        worksheet_players.set_column('C:C', 17, center_format)  # location, 3.5cm approx 13
        worksheet_players.set_column('D:D', 14, center_format)  # birthdate, 3.2cm approx 12
        worksheet_players.set_column('E:E', 12, center_format)  # rating

        worksheet_tournament = writer.sheets['Tournament']
        worksheet_tournament.set_column('A:A', 17)
        worksheet_tournament.set_column('B:B', 38)

    return players_df, tournament_df

parse_xml_and_save_to_excel('примерфайла.xml', 'example.xlsx')