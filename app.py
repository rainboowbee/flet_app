from flet import *
import example

path = ''
save_path = ''
def main(page: Page):
    page.title = 'XMLparser'

    def parse_xml(e):
        global path
        players_df, tournament_df = example.parse_xml_to_dataframe(path)

        players_rows = []
        for _, row in players_df.iterrows():
            cells = [
                DataCell(Text(str(row['№']))),
                DataCell(Text(row['ФИО'])),
                DataCell(Text(row['Город'])),
                DataCell(Text(row['Год рождения'])),
                DataCell(Text(row['рейтинг'])),
            ]
            players_rows.append(DataRow(cells=cells))
        
        data_players = DataTable(
            data_row_max_height=35,
            column_spacing=10,
            columns=[
                DataColumn(Text("№")),
                DataColumn(Text("ФИО")),
                DataColumn(Text("Город")),
                DataColumn(Text("Год рождения")),
                DataColumn(Text("рейтинг")),
            ],
            rows=players_rows,
        )

        rows_tournament = []
        for _, row in tournament_df.iterrows():
            for col in tournament_df.columns:
                cells = [
                    DataCell(Text(col)),
                    DataCell(Text(row[col]))
                ]
                rows_tournament.append(DataRow(cells=cells))
        
        tournament_columns = [
            DataColumn(Text("Info")),
            DataColumn(Text("Value"))
        ]
        
        tournament_data = DataTable(
            columns=tournament_columns,
            rows=rows_tournament
        )
        lv_players.controls.clear()
        lv_tournament.controls.clear()

        lv_players.controls.append(data_players)
        lv_tournament.controls.append(tournament_data)

        page.update()

    
    lv_players = ListView(expand=1)
    lv_tournament = ListView(expand=1)


    # Pick files dialog
    def pick_files_result(e: FilePickerResultEvent):
        global path
        path = ''
        if not e.files:
            selected_files.value = 'Файл не выбран'
        else:
            for el in e.files:
                selected_files.value = el.path
        path = selected_files.value
        parse_xml(path)

    pick_files_dialog = FilePicker(on_result=pick_files_result)
    selected_files = Text()

    # Save file dialog
    def save_file_result(e: FilePickerResultEvent):
        global save_path
        save_path = ''
        save_file_path.value = e.path + '.xlsx' if e.path else "Cancelled!"

        save_path = save_file_path.value
        example.parse_xml_and_save_to_excel(path, save_path)
        print('Успешно сохранено в ' + save_path)


    save_file_dialog = FilePicker(on_result=save_file_result)
    save_file_path = Text()

    def change_theme(e):
        if page.theme_mode == ThemeMode.LIGHT:
            page.theme_mode = ThemeMode.DARK
        else:
            page.theme_mode = ThemeMode.LIGHT


    # hide all dialogs in overlay
    page.overlay.extend([pick_files_dialog, save_file_dialog])


    page.add(
        Column(
            [
                Row(
                    [
                        ElevatedButton(
                            "Pick files",
                            icon=icons.UPLOAD_FILE,
                            on_click=lambda _: pick_files_dialog.pick_files(
                                allow_multiple=True
                                ),
                        ),
                        ElevatedButton(
                            "Save file",
                            icon=icons.SAVE,
                            on_click=lambda _: save_file_dialog.save_file(),
                        ),
                        IconButton(
                            icon=icons.AUTO_MODE,
                            on_click=change_theme
                        )
                    ]
                )
            ],
            adaptive=True
        ),
        Row(
            [
                Column(
                    [
                        Container(
                            content=lv_players,
                            width=700,
                            height=500,
                            border=border.all(2, colors.BLACK),
                        ),
                    ],
                ),
                Column(
                    [
                        Container(
                            content=lv_tournament,
                            width=700,
                            height=500,
                            border=border.all(2, colors.BLACK),
                        )
                    ],
                    adaptive=True
                )
            ],
            adaptive=True,
            alignment=MainAxisAlignment.CENTER
        ),
        Column(
            [
                Row(
                    [
                        IconButton(icon=icons.COPYRIGHT, url='https://github.com/rainboowbee'),
                        Text(
                            value='Made by Dmitry Glumov',

                        ),
                    ],
                    alignment=MainAxisAlignment.CENTER,
                    adaptive=True
                )
            ],
            alignment=MainAxisAlignment.END,
            adaptive=True
        )
    )

app(target=main)

