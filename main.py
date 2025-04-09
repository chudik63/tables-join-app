import flet as ft
from pages import Home
from data import Data

def main(page: ft.Page):
    page.window_width = 350
    page.window_height = 700

    data = Data()

    def route_change(e):
        if page.route == "/home":
            page.views.clear()
            page.views.append(Home(page, data))
        page.update()

    page.on_route_change = route_change

    page.go('/home')

if __name__ == "__main__":
    ft.app(target=main)