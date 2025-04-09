import flet as ft

class Home(ft.View):
    def __init__(self, page: ft.Page, data):
        super().__init__()
        self.page = page
        self.data = data
        self.file_uploader = ft.FilePicker(on_result=self.on_upload_result, data='data')
        self.file_saver = ft.FilePicker(on_result=self.on_save_result)
        self.page.overlay.append(self.file_uploader)
        self.page.overlay.append(self.file_saver)

        self.dropdown = ft.Dropdown(
            label="Месяц",
            on_change=self.dropdown_changed,
            options=[
                ft.dropdown.Option("Январь"),
                ft.dropdown.Option("Февраль"),
                ft.dropdown.Option("Март"),
                ft.dropdown.Option("Апрель"),
                ft.dropdown.Option("Май"),
                ft.dropdown.Option("Июнь"),
                ft.dropdown.Option("Июль"),
                ft.dropdown.Option("Август"),
                ft.dropdown.Option("Сентябрь"),
                ft.dropdown.Option("Октябрь"),
                ft.dropdown.Option("Ноябрь"),
                ft.dropdown.Option("Декабрь"),
            ],
            autofocus=True,
            width = 280
        )

        self.data_button_clicked = False
        self.source_button_clicked = False
        self.dropdown_checked = False

        self.result_button = ft.ElevatedButton('Выполнить вставку', on_click=self.insert_table, width = 280, disabled = True)

        self.dialog = ft.AlertDialog(
            content=ft.Container(
                content=ft.ProgressRing(width=50, height=50, stroke_width = 3),
                alignment=ft.alignment.center,
                width = 60,
                height = 60
            ),
            alignment=ft.alignment.center,
            modal = True,
        )

        self.controls = [
            ft.Container(
                content=ft.Column(
                        [   
                            ft.ElevatedButton('Загрузить ПЖРТ', on_click=self.upload_data, width = 280),
                            self.dropdown,
                            ft.ElevatedButton('Загрузить 415', on_click=self.upload_source, width = 280),
                            self.result_button,
                        ],
                        alignment=ft.MainAxisAlignment.CENTER,
                        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                        expand=True,
                    ),
                alignment=ft.alignment.center,
                expand=True,
            )
        ]
    
    def dropdown_changed(self, e):
        self.data.set_month(self.dropdown.value)
        self.result_button.disabled = not(self.source_button_clicked and self.data_button_clicked and self.dropdown.value)
        self.page.update()
    
    def upload_data(self, e):
        self.file_uploader.data = 'data'
        self.file_uploader.pick_files(allow_multiple=False)

    def upload_source(self, e):
        self.file_uploader.data = 'source'
        self.file_uploader.pick_files(allow_multiple=False)

    def on_upload_result(self, e):
        if self.file_uploader.data == 'data' and e.files:
            self.data.read_data_table(e.files[0].path)
            self.data_button_clicked = True
        elif self.file_uploader.data == 'source' and e.files:
            self.data.read_source_table(e.files[0].path)
            self.source_button_clicked = True

        self.result_button.disabled = not(self.source_button_clicked and self.data_button_clicked and self.dropdown.value)
        self.page.update()
    
    def open_dialog(self):
        self.page.dialog = self.dialog
        self.dialog.open = True
        self.page.update()

    def close_ring(self):
        self.dialog.title = ft.Text("Таблица сохранена")
        self.dialog.content = None
        self.dialog.actions = [
            ft.TextButton("Выход", on_click=lambda _: self.page.window_close()),
        ]
        self.dialog.actions_alignment=ft.MainAxisAlignment.END
        self.page.update()
    
    def on_save_result(self, e):
        self.open_dialog()
        self.data.get_result_table(e.path)
        self.close_ring()
        

    def insert_table(self, e):
        self.file_saver.get_directory_path()