'''
V.1.2.0

Цветовая схема https://colorscheme.ru/#3uk1E4ih8xnI-

'''
import os
from docx import Document
import pandas as pd
import re
from datetime import datetime, timedelta, date
from kivy.clock import Clock
from openpyxl import load_workbook
from docx.enum.text import WD_ALIGN_PARAGRAPH
from kivy.config import Config
from kivy.app import App
from kivymd.app import MDApp
from kivymd.uix.screen import MDScreen
from kivymd.uix.screenmanager import ScreenManager
from kivy.lang import Builder
from kivy.uix.popup import Popup
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivymd.uix.fitimage import FitImage
from kivymd.uix.button import MDButton
from kivymd.icon_definitions import md_icons
from kivymd.uix.pickers import MDModalInputDatePicker
from kivymd.uix.appbar import MDTopAppBar
from kivymd.uix.appbar import MDActionTopAppBarButton
from kivymd.uix.menu import MDDropdownMenu
from kivymd.uix.tooltip import MDTooltip

from kivy.core.window import Window

os.environ['KIVY_IMAGE'] = 'pil'

pd.set_option('display.max_columns', 8)
pd.set_option('display.max_rows', 10)

KV = """
MainScreen:
    id: main
    md_bg_color: 'C4CCCC'
   
    FitImage:
        id: bck_image
        size_hint_x: 1
        size_hint_y: 1
        source: 'data/bkg.png'
        opacity: .4

        
    MDBoxLayout:
        id: top_box
        orientation: 'vertical'

        MDTopAppBar:
            type: 'small'
            theme_bg_color: 'Custom'
            md_bg_color: 'BFC6D1'

            
            MDTopAppBarLeadingButtonContainer:
                spacing: 5
                id: btn_cnt_1


                ButtonToolTip:
                    id: add_to_db
                    icon: 'package-down'
                    theme_bg_color: 'Custom'
                    md_bg_color: 'C4CCCC'
                    on_release: root.output_label('docs')
                    on_press: root.view.open()

                    MDTooltipPlain:
                        text: 'Добавить находящиеся в папке "Docs" \\n' \
                        'файлы досье в общую БД'

                
            
                
            MDTopAppBarTrailingButtonContainer:
                spacing: 5
                id: btn_cnt_2
                
                ButtonToolTip:
                    id: choose_dates
                    icon: 'calendar-range'
                    theme_bg_color: 'Custom'
                    md_bg_color: 'C4CCCC'
                    on_release: root.show_datepicker()

                    MDTooltipPlain:
                        text: 'Задать даты выборки досье \\n' \
                        'для обработки и предпросмотра данных'

                ButtonToolTip:
                    id: view_db
                    icon: 'view-list'
                    theme_bg_color: 'Custom'
                    md_bg_color: 'C4CCCC'
                    on_release: root.open_menu(self)

                    MDTooltipPlain:
                        text: 'Доступные серии препарата \\n' \
                        'в выбранном диапазоне'
                
                ButtonToolTip:
                    id: report
                    icon: 'file-export'
                    theme_bg_color: 'Custom'
                    md_bg_color: 'C4CCCC'
                    on_press: root.articles_report()

                    MDTooltipPlain:
                        text: 'Сформировать отчет'
            
                                    
            MDTopAppBarTitle:
                markup: True
                text: 'База данных досье РФЛП [sup]18[/sup]F-ФДГ'
                theme_font_name: 'Custom'
                font_name: 'CENTURY'
                text_color: '14213D'
                #pos_hint: {"center_x": .5}

        MDBoxLayout:
            id: bottom_box
            orientation: 'vertical'
                    
        
                
"""


class ButtonToolTip(MDTooltip, MDActionTopAppBarButton):
    pass
        
        
class MainScreen(MDScreen):

    def __init__(self, *args, **kwargs):
        super().__init__(**kwargs)
        self.all_data = {}
        # загрузка текущей базы данных по досье
        self.full_data = pd.read_csv('full_data.csv', index_col=0)
##        self.full_data = pd.DataFrame()

        # диапазон выборки досье
        self.daterange = [date(1678,1,1), date.today() + timedelta(days=1)]
        
        # всплывающее окно во время загрузки файлов
        self.view = Popup(title='Работаем',
                          title_color=[.3, .3, .3, 1],
                          content=Label(text='Придется немного подождать'),
                          separator_color=[.14, .35, .48, 1],
                          size_hint=(None, None),
                          size=(300, 150),
                          auto_dismiss=True,
                          overlay_color=[0, 0, 0, .5],
                          background='',
                          background_color=[.65, .72, .77, 1]
                          )

        self.full_data_view = Popup(title='База данных досье',
                          title_color=[.3, .3, .3, 1],
                          content=Label(text=str(self.full_data.head(10))),
                          separator_color=[.14, .35, .48, 1],
                          size_hint=(None, None),
                          auto_dismiss=True,
                          overlay_color=[0, 0, 0, .5],
                          background='',
                          background_color=[.65, .72, .77, 1]
                          )

        
    def open_menu(self, item):
        unique_series_full_data = self.full_data.groupby(['Серия препарата', 'Дата'],
                                             as_index=False). \
                                      agg(lambda x: list(x.unique())). \
                                      sort_values(['Дата'])
        
        print(unique_series_full_data)
        unique_series_full_data['Дата'] = pd.to_datetime(
            unique_series_full_data['Дата'])
        usfd_selection_by_date = unique_series_full_data.query(
            '@self.daterange[0] <= Дата < @self.daterange[1]')
        print(usfd_selection_by_date)
        menu_items = [
            {
                "text": f"{i}",
                "on_release": lambda x=f"{i}": self.menu_callback(x),
            } for i in usfd_selection_by_date['Серия препарата']
        ]
        MDDropdownMenu(caller=item, items=menu_items).open()

    def menu_callback(self, text_item):
        print(text_item)
        series_prep = text_item
                
    def show_datepicker(self):
        '''Окно выбора диапазон дат отчета
        Вызывается нажатием кнопки 'выбрать дату отчета'
        
        '''
        datepicker = MDModalInputDatePicker(mode='range',
                                            supporting_input_text='Укажите даты',
                                            supporting_text=' ',
                                            text_button_cancel='Отмена',
                                            text_button_ok='Принять',
                                            date_format='dd/mm/yyyy',
                                            )
        datepicker.bind(on_ok=self.on_ok)
        datepicker.bind(on_cancel=self.on_cancel)
        datepicker.open()
        
    def on_ok(self, instance_datepicker):
        '''Вызывается нажатием кнопки "Ок" в окне выбора дат
        При указании начальной и конечной даты обновляет значения переменной
        self.daterange в формате [datetime.date(yyy, mm, dd), ...]
        Закрывает окно datepicker
        
        '''
        
        if len(instance_datepicker.get_date()) > 1:
            print(instance_datepicker.get_date())
            self.daterange = sorted(instance_datepicker.get_date())
            print(self.daterange, ' self.daterange')
            return instance_datepicker.dismiss()
        else:
            print('Enter dates')

    def on_cancel(self, instance_datepicker):
        return instance_datepicker.dismiss()

    def source_datas(self, folder):
        '''

        '''
        docslist = []
        series = pd.Series(dtype='string')
        
        for root, dirs, files in os.walk(os.path.abspath(folder)):
            # список всех файлов
            for file in files:
                docslist.append(os.path.join(root, file))
        
        for file in docslist:
            try:
                # ошибка, если формат файла не является excel
                # чтение таблицы с листа Досье, стр. 7
                excelDoc = pd.read_excel(file,
                                         sheet_name=3,
                                         usecols='D, Y, AH, AT, AZ',
                                         header=0,
                                         skiprows=205,
                                         nrows=19, #для тонкой настройки можно добавить переменную
                                         engine='openpyxl')\
                                         .dropna().astype('str')
                # вычленение из названия файлов номера серии
                s = re.search(r'S\d{11}', file)[0]
                # определение даты из номера серии
                date = datetime.strptime(s[-6:], '%d%m%y').date()
                
                # добавление столбцов с номером серии препарата и даты
                excelDoc['Серия препарата'] = pd.Series([s])
                excelDoc['Серия препарата'] = excelDoc['Серия препарата'].fillna(s)
                excelDoc['Дата'] = pd.Series(date)
                excelDoc['Дата'] = excelDoc['Дата'].fillna(date).astype('str')
                # добавление в общую базу новых данных с прочитанных досье
                self.full_data = pd.concat([self.full_data, excelDoc])
                self.full_data.drop_duplicates(keep='last', inplace=True)
                
            except Exception as e:
                print(e)
                
        print(self.full_data)
        self.full_data.to_csv('full_data.csv')
        
        return


    def output_label(self, folder):
        '''Формирует из полученного словаря текстовые строки
        Обновляет текст в центральном поле приложения (ids['label'])
        на полученные строки

        Args:
            folder (str): название папки с файлами

        Calls:
            foo checks(folder): принимает возвращаемое значение
                                функции в качестве переменной
        Returns:
            None
        
        '''
        # закрываем открытое окно popup
        self.view.dismiss()
        # 
        self.source_datas(folder)

    
    def name_report(self):
        '''Вызывается нажатием кнопки 'Отчет'
        Popup-окно с полем для ввода названия сохраняемого файла
        
        Calls:
           foo report(): После подтверждения ввода текста (enter)

        '''
        self.p_report = Popup(title='Название отчета',
                              title_color=[.3, .3, .3, 1],
                              content=TextInput(multiline=False,
                                                on_text_validate=self.report,
                                                size_hint=(None, None),
                                                size=(250, 30)
                                                ),
                              separator_color=[.14, .35, .48, 1],
                              size_hint=(None, None),
                              size=(300, 100),
                              auto_dismiss=False,
                              overlay_color=[0, 0, 0, .5],
                              background='',
                              background_color=[.65, .72, .77, 1]
                              )
        self.p_report.open()

    def articles_report(self):
        '''articles_report(self)
        Создает выборку  из БД временного диапазона из self.daterange
        Группирует БД по наименованию товара.

        Returns:
            pd.DataFrame()
                ДФ содержит столбцы 'Наименование сырья/материала' и
                'Номер серии'(список артикулов)

        '''
        # Изменим тип данных столбца Дата в date
        full_data_to_dates = self.full_data
        full_data_to_dates['Дата'] = pd.to_datetime(full_data_to_dates['Дата'])
        # выборка БД с учетом дат, сгруппированная по наименованию сырья
        group_data = full_data_to_dates.query(
            '@self.daterange[0] <= Дата < @self.daterange[1]'). \
            drop('Дата', axis=1). \
            groupby('Наименование сырья/материала').\
            agg(lambda x: x.unique()).reset_index()
        # Выводит БД Наименование сырья - все уникальные номера серий
        articles_data = pd.concat([group_data['Наименование сырья/материала'].to_frame(), 
                                  group_data['Номер серии'].to_frame()],
                                  axis=1
                                 )
        print(articles_data)
        # сохранение БД с номерами серий
        articles_data.to_csv('articles_data.csv')
            
        return articles_data
        
class SourceMaterialsCheckApp(MDApp):
    def build(self):
        self.icon = 'data/smcicon.png'
        sm = ScreenManager()
        sm.add_widget(MainScreen(name='main_screen'))
        return Builder.load_string(KV)


if __name__ == "__main__":                    
    SourceMaterialsCheckApp().run()
