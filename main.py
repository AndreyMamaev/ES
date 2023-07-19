import pandas as pd
import PySimpleGUI as sg
import numpy as np
from datetime import datetime, timedelta
from typing import List

sg.theme('DarkTeal12')


class Window():
    REQUIRED_COLUMNS = {
        'Номер карточки',
        'ТБ/ЦА',
        'Подразделение',
        'Статус КЗОП',
        'Вид события',
        'Потерпевший СБЕР',
        'Дата подачи ЗОП',
        'Дата возбуждения УД',
        'Дата передачи дела в суд первой инстанции',
        'Подозреваемые',
        'Сумма ущерба',
        'Ущерб возмещенный',
    }
    DATE_FORMAT = '%d.%m.%Y'
    TB = {
        'Байкальский банк (ББ)': [
            '-',
            'Бурятское ГОСБ №8601',
            'Читинские ГОСБ №8600',
            'Якутское ГОСБ №8603',
        ],
        'Волго-Вятский банк (ВВБ)': [
            '-',
            'Владимирское ГОСБ №8611',
            'Кировское ГОСБ №8612',
            'Марий Эл ГОСБ №8614',
            'Мордовское ГОСБ №8589',
            'Пермское ГОСБ №6984',
            'Татарстан ГОСБ №8610',
            'Удмуртское ГОСБ №8618',
            'Чувашское ГОСБ №8613',
        ],
        'Дальневосточный банк (ДВБ)': [
            '-',
            'Биробиджанский ГОСБ №4157',
            'Благовещенский ГОСБ №8636',
            'Головное отделение по Хабаровскому краю',
            'Камчатский ГОСБ №8556',
            'Приморский ГОСБ №8635',
            'Северо-Восточный ГОСБ №8645',
            'Чукотское головное отделение  ',
            'Южно-Сахалинский ГОСБ №8567',
        ],
        'Московский банк (МБ)': [
            '-',
        ],
        'Поволжский банк (ПБ)': [
            '-',
            'Астраханское ГОСБ №8625',
            'Волгоградское ГОСБ №8621',
            'Оренбургское ГОСБ №8623',
            'Пензенское ГОСБ №8624',
            'Саратовское отделение №8622',
            'Ульяновское ГОСБ №8588',
        ],
        'Сибирский банк (СБ)': [
            '-',
            'Алтайское ГОСБ №8644',
            'Кемеровского ГОСБ №8615',
            'Красноярское отделение №8646',
            'Новосибирское ГОСБ №8047',
            'Омское ГОСБ №8634',
            'Томское отделение №8616',
        ],
        'Северо-Западный банк (СЗБ)': [
            '-',
            'Архангельское отделение №8637',
            'Вологодское отделение №8638',
            'Калининградское отделение №8626',
            'Карельское отделение №8628',
            'Коми отделение №8617',
            'Мурманское отделение №8627',
            'Новгородское отделение №8629',
            'Псковское отделение №8630',
        ],
        'Среднерусский банк (СРБ)': [
            '-',
            'Брянское ГОСБ №8605',
            'Восточное Головное отделение',
            'Западное Головное отделение',
            'Ивановское ГОСБ №8639',
            'Калужское ГОСБ №8608',
            'Костромское ГОСБ №8640',
            'Рязанское ГОСБ №8606',
            'Северное Головное отделение',
            'Смоленское ГОСБ №8609',
            'Тверское ГОСБ №8607',
            'Тульское ГОСБ №8604',
            'Южное Головное отделение ',
            'Ярославское ГОСБ №0017',
        ],
        'Уральский банк (УБ)': [
            '-',
            'Башкирское ГОСБ №8598',
            'Западно-Сибирское ГОСБ №8647',
            'Курганское ГОСБ №8599',
            'Новоуренгойское ГОСБ №8369',
            'Сургутское ГОСБ №5940',
            'Челябинское ГОСБ №8597',
        ],
        'Центрально-Черноземный банк (ЦЧБ)': [
            '-',
            'Белгородское ГОСБ №8592',
            'Курское ГОСБ № №8596',
            'Липецкое ГОСБ №8593',
            'Орловское ГОСБ №8595',
            'Тамбовское ГОСБ №8594',
        ],
        'Юго-Западный банк (ЮЗБ)': [
            'Дагестанское ГОСБ №8590',
            'Ингушское ГОСБ №8633',
            'Кабардино-Балкарское ГОСБ №8631',
            'Калмыцкое ГОСБ №8579',
            'Карачаево-Черкесское ГОСБ №8585',
            'Краснодарское ГОСБ №8619',
            'Ростовское ГОСБ №5221',
            'Северо-Осетинское ГОСБ №8632',
            'Ставропольское ГОСБ №5230',
            'Чеченское ГОСБ №8643',
        ],
    }
    TB_REPLACE = {
        'СиБ': 'СБ',
        'ПВБ': 'ПБ'
    }

    def __init__(self):
        self.individ_report = pd.DataFrame(
            columns=[
                'Кпкуо', 'Кэр', 'Кэву', 'Кэппл', 'Кмпк',
                'Кол-во исключений',
                'Кпкуо/Кэр: Кол-во ВУД', 'Кпкуо/Кэр: Передано в суд',
                'Кпкуо/Кэр: Установлено лиц',
                'Кэву: Сумма прич.', 'Кэву: Сумма возм.',
                'Кэппл: Кол-во ВУД', 'Кэппл: Передано в суд(>365)',
                'Кэппл: Передано в суд(<=365)',
                'Кмпк: Кол-во заяв.', 'Кмпк: Кол-во ВУД',
            ],
        )
        self.legal_registry_tables = {
            'legal_kvu_kur': {
                'name': 'КВУ_КУР',
                'df': pd.DataFrame(
                    columns=[
                        'КВУ_Наименование',
                        'КВУ_Номер карточки',
                        'КВУ_ТБ',
                    ],
                )
            },
            'legal_kvu_kvupp_kpupp': {
                'name': 'КВУ_КВУПП_КПУПП',
                'df': pd.DataFrame(
                    columns=[
                        'КВУ_Наименование',
                        'КВУ_Номер карточки',
                        'КВУ_ТБ',
                        'Подразделение',
                        'КПУПП_Ущерб причиненный',
                        'КВУПП_Ущерб возмещеннный',
                    ],
                )
            },
            'legal_koup_kur': {
                'name': 'КОУП_КУР',
                'df': pd.DataFrame(
                    columns=[
                        'КОУП_Наименование',
                        'КОУП_Номер карточки',
                        'КОУП_ТБ',
                    ],
                )
            },
            'legal_koup_zpp_vpp': {
                'name': 'КОУП_ЗПП_ВПП',
                'df': pd.DataFrame(
                    columns=[
                        'КОУП_Наименование',
                        'КОУП_Номер карточки',
                        'КОУП_ТБ',
                        'Подразделение',
                    ],
                )
            },
        }
        self.legal_report = pd.DataFrame(
            columns=[
                'Коуп', 'Кву',
                'Коуп: Кол-во', 'Коуп: Кол-во ВУД',
                'Коуп: Кол-во искл.', 'Коуп: Кол-во доб.',
                'Кву: Сумма', 'Кву: Кол-во',
                'Кву: Кол-во искл.', 'Кву: Кол-во доб.',
            ]
        )
        self.__tabs = []
        self.__open_files()
        self.__get_dfs()
        if self.__check_dfs():
            self.__create_window()

    def __open_files(self):
        """
        Вызывает окно выбора файлов.
        Запысывает пути к выбранным файлам в self.files
        """
        self.files = sg.popup_get_file(
            'Select a file',
            title="File selector",
            multiple_files=True
        ).split(';')

    def __get_dfs(self):
        """
        Cоздает словарь self.dfs с элементами типа
        Имя файла: Dataframe из файла
        """
        self.dfs = {
            file_name: pd.read_excel(file_name) for file_name in self.files
        }

    def __check_dfs(self) -> bool:
        """
        Проверяет значения словаря self.dfs на наличие
        в df обязательных столбцов.
        В случае отсутсвия хотя бы одного столбца
        вызывает всплывающее окно и прекращает выполнение программы.
        """
        for file_name, df in self.dfs.items():
            diff = self.REQUIRED_COLUMNS.difference(df.columns.values)
            if diff != set():
                sg.popup(
                    f'В файле {file_name} отсутствуют столбы:\n'
                    f'{chr(10).join(diff)}\n'
                    f'Проверьте файл и повторите попытку.'
                )
                return False
        return True

    def __create_window(self):
        """
        Создает рабочее окно программы.
        """
        self.layout = [
            [
                sg.TabGroup([[
                    sg.Tab(
                        filename,
                        [[sg.Table(
                            values=df.values.tolist(),
                            headings=df.columns.values.tolist(),
                            vertical_scroll_only=False,
                            justification='center',
                            alternating_row_color=sg.theme_button_color()[1],
                            selected_row_colors='white on black',
                        )]]
                    ) for filename, df in self.dfs.items()
                ]])
            ],
        ]
        self.__create_tabs()
        self.window = sg.Window(
            'ES', self.layout,
            resizable=True,
            finalize=True,
        )
        self.window.Maximize()
        self.__window_loop()

    def __create_tabs(self):
        self.__create_individ_tab()
        self.__create_legal_tab()
        self.layout += [
            [sg.TabGroup([self.__tabs], expand_x=True, expand_y=True,)]
        ]

    def __create_individ_tab(self):
        self.__tabs.append(sg.Tab('Физические лица', [
            [
                sg.Column([
                    [
                        sg.Text(report, size=(10, 1)),
                        sg.Input(
                            key=f'individ_{key}_date_start',
                            disabled=True,
                            size=(10, 1),
                        ),
                        sg.CalendarButton(
                            "Начало периода",
                            close_when_date_chosen=False,
                            target=f'individ_{key}_date_start',
                            format=self.DATE_FORMAT,
                            size=(12, 1),
                        ),
                        sg.Input(
                            key=f'individ_{key}_date_finish',
                            disabled=True,
                            size=(10, 1)
                        ),
                        sg.CalendarButton(
                            "Конец периода",
                            close_when_date_chosen=False,
                            target=f'individ_{key}_date_finish',
                            format=self.DATE_FORMAT,
                            size=(12, 1)
                        ),
                    ] for report, key in (
                        ('Кпкуо/Кэр', 'kpkuo_ker'),
                        ('Кэву', 'kevu'),
                        ('Кэппл', 'keppl'),
                        ('Кмпк', 'kmpk'),
                    )
                ], element_justification='left',),
                sg.Column([
                    [
                        sg.Text('Исключения:',),
                        sg.Input(
                            key='individ_exception_file',
                            visible=False,
                            enable_events=True,
                        ),
                        sg.FileBrowse(
                            'Выбрать из файла',
                            target='individ_exception_file',
                            file_types=(('Excel', '.xlsx'),),
                        )
                    ],
                    [sg.Multiline(
                        key='individ_exceptions',
                        expand_x=True,
                        expand_y=True,
                        size=(30, 5),
                    )],
                ], element_justification='right',)
            ],
            [sg.Text('Результаты')],
            [sg.Table(
                values=[],
                headings=(['ТБ/ГОСБ'] + self.individ_report.columns.tolist()),
                key='individ_report',
                expand_x=True,
                expand_y=True,
                vertical_scroll_only=False,
                col_widths=[20] + 5 * [5] + 11 * [15],
                auto_size_columns=False,
                justification='center',
                alternating_row_color=sg.theme_button_color()[1],
                selected_row_colors='white on black',
            )],
            [
                sg.Button(
                    'Расчитать',
                    key='individ_calc_report'
                ),
                sg.Input(
                    visible=False,
                    enable_events=True,
                    key='individ_save_path'
                ),
                sg.FileSaveAs(
                    'Сохранить как',
                    key='individ_save_button',
                    file_types=(('Excel', '.xlsx'),),
                )
            ],
        ]))

    def __create_legal_tab(self):
        self.__tabs.append(sg.Tab(
            'Юридические лица',
            [
                [
                    sg.Text(
                        "Коуп/Кву",
                        size=(10, 1)
                    ),
                    sg.Input(
                        key='legal_koup_kvu_date_start',
                        disabled=True,
                        size=(10, 1),
                    ),
                    sg.CalendarButton(
                        "Начало периода",
                        close_when_date_chosen=False,
                        target='legal_koup_kvu_date_start',
                        format=self.DATE_FORMAT,
                        size=(12, 1),
                    ),
                    sg.Input(
                        key='legal_koup_kvu_date_finish',
                        disabled=True,
                        size=(10, 1)
                    ),
                    sg.CalendarButton(
                        "Конец периода",
                        close_when_date_chosen=False,
                        target='legal_koup_kvu_date_finish',
                        format=self.DATE_FORMAT,
                        size=(12, 1)
                    ),
                    sg.Input(
                        visible=False,
                        key='legal_koup_kvu_files',
                        enable_events=True,
                    ),
                    sg.FilesBrowse(
                        'Загрузить реестры',
                        target='legal_koup_kvu_files',
                    )
                ],
                [
                    sg.TabGroup([[
                        sg.Tab(
                            reg['name'],
                            [[sg.Table(
                                values=[],
                                headings=reg['df'].columns.tolist(),
                                key=key,
                                expand_x=True,
                                justification='center',
                                alternating_row_color=sg.theme_button_color()[1],
                                selected_row_colors='white on black',
                                enable_click_events=True,
                            )], [sg.Button(
                                'Добавить запись',
                                key=f'{key}_add',
                                enable_events=True
                            )]]
                        ) for key, reg in self.legal_registry_tables.items()
                    ]], expand_x=True,)
                ],
                [
                    sg.Table(
                        values=[],
                        headings=(
                            ['ТБ/ГОСБ'] + self.legal_report.columns.tolist()
                        ),
                        key='legal_report',
                        expand_x=True,
                        expand_y=True,
                        vertical_scroll_only=False,
                        col_widths=[20] + 2 * [5] + 8 * [15],
                        auto_size_columns=False,
                        justification='center',
                        alternating_row_color=sg.theme_button_color()[1],
                        selected_row_colors='white on black',
                    )
                ],
                [
                    sg.Button('Расчитать', key='legal_calc_report'),
                    sg.Input(
                        visible=False,
                        enable_events=True,
                        key='legal_save_path'
                    ),
                    sg.FileSaveAs(
                        'Сохранить как',
                        key='legal_save_button',
                        file_types=(('Excel', '.xlsx'),),
                    )
                ]
            ]
        ))

    def __check_individ_events(self, event, values):
        match event:
            case 'individ_exception_file':
                exceptions = '\n'.join(pd.read_excel(
                    values['individ_exception_file'],
                    header=None
                )[0].values)
                if values['individ_exceptions'] != '':
                    values['individ_exceptions'] += '\n'
                self.window['individ_exceptions'].update(
                    values['individ_exceptions'] + exceptions
                )
            case 'individ_save_path':
                self.individ_report.to_excel(values['individ_save_path'])
            case 'individ_calc_report':
                if any([
                    values[v] == '' for v in values if str(v).endswith((
                        '_start', '_finish'
                    )) and str(v).startswith('individ_')
                ]):
                    sg.popup('Заполните даты расчета коэффициентов')
                else:
                    for file_name, df in self.dfs.items():
                        self.individ_report = pd.DataFrame()
                        report = ES_individ(
                            df=df,
                            exceptions=values['individ_exceptions'].split('\n'),
                            kpkuo_ker_date_start=datetime.strptime(
                                values['individ_kpkuo_ker_date_start'],
                                self.DATE_FORMAT
                            ).date(),
                            kpkuo_ker_date_finish=datetime.strptime(
                                values['individ_kpkuo_ker_date_finish'],
                                self.DATE_FORMAT
                            ).date(),
                            kmpk_date_start=datetime.strptime(
                                values['individ_kmpk_date_start'],
                                self.DATE_FORMAT
                            ).date(),
                            kmpk_date_finish=datetime.strptime(
                                values['individ_kmpk_date_finish'],
                                self.DATE_FORMAT
                            ).date(),
                            keppl_date_start=datetime.strptime(
                                values['individ_keppl_date_start'],
                                self.DATE_FORMAT
                            ).date(),
                            keppl_date_finish=datetime.strptime(
                                values['individ_keppl_date_finish'],
                                self.DATE_FORMAT
                            ).date(),
                            kevu_date_start=datetime.strptime(
                                values['individ_kevu_date_start'],
                                self.DATE_FORMAT
                            ).date(),
                            kevu_date_finish=datetime.strptime(
                                values['individ_kevu_date_finish'],
                                self.DATE_FORMAT
                            ).date(),
                        )
                        self.individ_report = pd.concat([
                            self.individ_report, report.report
                        ])
                    self.window['individ_report'].update(
                        [[index] + values for index, values in zip(
                            self.individ_report.index.values.tolist(),
                            self.individ_report.values.tolist()
                        )]
                    )

    def __check_legal_events(self, event, values):
        if event == 'legal_koup_kvu_files':
            registries = values['legal_koup_kvu_files'].split(';')
            registry_names = set(
                map(
                    lambda x: x + '.xlsx',
                    [reg['name'] for reg in self.legal_registry_tables.values()]
                )
            )
            if registry_names.difference([registry.split('/')[-1] for registry in registries]) != set():
                sg.popup(
                    f'Загрузите 4 файла с названиями\n'
                    f'{chr(10).join(registry_names)}'
                )
                self.window['legal_koup_kvu_files'].update('')
            else:
                for key, reg in self.legal_registry_tables.items():
                    columns = reg['df'].columns.tolist()
                    registry_file = [registry for registry in registries if reg['name'] + '.xlsx' in registry][0]
                    new_df = pd.read_excel(registry_file)
                    if 'Подразделение' in reg['df'].columns.values:
                        new_df['Подразделение'] = ''
                    for column in new_df.columns.values:
                        new_df.rename(
                            columns={column: column.strip()},
                            inplace=True
                        )
                        if column.endswith('_ТБ'):
                            for i, value in new_df[column].items():
                                value = value.strip()
                                if value in self.TB_REPLACE.keys():
                                    value = self.TB_REPLACE[value]
                                for tb in self.TB.keys():
                                    if f'({value})' in tb:
                                        new_df.loc[i, column] = tb
                    self.legal_registry_tables[key]['df'] = new_df[columns]
                    self.window[key].update(
                        values=self.legal_registry_tables[key]['df'].values.tolist()
                    )
        elif isinstance(event, tuple) and event[1] == '+CLICKED+':
            if (
                event[2][0] not in (None, -1) and
                self.legal_registry_tables[event[0]]['df'].columns[event[2][1]] == 'Подразделение'
            ):
                for key, reg in self.legal_registry_tables.items():
                    if event[0] == key:
                        tb = reg['df'].loc[
                            event[2][0],
                            [column for column in reg['df'].columns if column.endswith('_ТБ')][0]
                        ]
                        break
                win = sg.Window(
                    'Выберите подразделение',
                    [
                        [
                            sg.Combo(
                                self.TB[tb],
                                key='combo',
                                expand_x=True,
                                readonly=True
                            )
                        ],
                        [
                            sg.Button('OK', key='ok'),
                            sg.Button('Cancel', key='cancel')
                        ]
                    ],
                    size=(400, 70)
                )
                e, v = win.read()
                if e == 'ok':
                    self.legal_registry_tables[key]['df'].loc[
                        event[2][0],
                        'Подразделение'
                    ] = v['combo']
                    self.window[key].update(
                        values=self.legal_registry_tables[key]['df'].values.tolist()
                    )
                win.close()
        elif event.endswith('_add'):
            key = event.replace('_add', '')
            df = self.legal_registry_tables[key]['df']
            win = sg.Window(
                'Добавить запись',
                [
                    [
                        sg.Text(
                            'Заполните обязательные поля!',
                            visible=False,
                            key='warn'
                        )
                    ],
                    [
                        sg.Column(
                            [
                                [sg.Text(column)],
                                [
                                    sg.Input(size=(len(column) + 10, 1))
                                    if not column.endswith((
                                        '_ТБ',
                                        'Подразделение'
                                    ))
                                    else (
                                        sg.Combo(
                                            list(self.TB.keys()),
                                            enable_events=True,
                                            key='tb',
                                            readonly=True
                                        )
                                        if column != 'Подразделение'
                                        else sg.Combo(
                                            [],
                                            disabled=True,
                                            key='gosb',
                                            size=(len(column) + 10, 1),
                                            readonly=True
                                        )
                                    )
                                ]
                            ]
                        ) for column in df.columns.values
                    ],
                    [
                        sg.Button('OK', key='ok'),
                        sg.Button('Cancel', key='cancel')
                    ]
                ]
            )
            while True:
                e, v = win.read()
                if e in (sg.WIN_CLOSED, 'cancel'):
                    break
                elif e == 'tb':
                    if 'gosb' in v.keys():
                        win['gosb'].update(
                            values=self.TB[v['tb']],
                            disabled=False
                        )
                elif e == 'ok':
                    if '' in list(v.values())[1:]:
                        win['warn'].update(visible=True)
                    else:
                        self.legal_registry_tables[key]['df'].loc[len(df.index)] = list(v.values())
                        self.window[key].update(
                            values=self.legal_registry_tables[key]['df'].values.tolist()
                        )
                        break
            win.close()
        elif event == 'legal_save_path':
            self.legal_report.to_excel(values['legal_save_path'])
        elif event == 'legal_calc_report':
            if '' in (
                values['legal_koup_kvu_date_start'],
                values['legal_koup_kvu_date_finish']
            ):
                sg.popup('Заполните даты расчета коэффициентов')
            elif any(['' in reg['df']['Подразделение'].values for reg in self.legal_registry_tables.values() if 'Подразделение' in reg['df'].columns.values]):
                sg.popup(
                    'Заполните значения подразделений в таблицах реестров'
                )
            else:
                self.legal_report = pd.DataFrame()
                for file_name, df in self.dfs.items():
                    report = ES_legal(
                        df,
                        registries={
                            key: reg['df'] for key, reg in self.legal_registry_tables.items()
                        },
                        koup_kvu_date_start=datetime.strptime(
                            values['legal_koup_kvu_date_start'],
                            self.DATE_FORMAT
                        ).date(),
                        koup_kvu_date_finish=datetime.strptime(
                            values['legal_koup_kvu_date_finish'],
                            self.DATE_FORMAT
                        ).date(),
                    )
                    self.legal_report = pd.concat([
                        self.legal_report,
                        report.report
                    ])
                self.window['legal_report'].update(
                    [[index] + values for index, values in zip(
                        self.legal_report.index.values.tolist(),
                        self.legal_report.values.tolist()
                    )]
                )

    def __window_loop(self):
        while True:
            event, values = self.window.read()
            if event == sg.WIN_CLOSED:
                break
            elif (
                isinstance(event, str) and
                event.startswith('individ_')
            ):
                self.__check_individ_events(event, values)
            elif (
                (
                    isinstance(event, tuple) and
                    event[0].startswith('legal_')
                ) or
                event.startswith('legal_')
            ):
                self.__check_legal_events(event, values)


class ES_individ():
    GROUP_GOSB = {
        'Среднерусский банк (СРБ)': {
            '[Аппарат СРБ]': [
                'Южное Головное отделение',
                'Восточное Головное отделение',
                'Северное Головное отделение',
                'Западное Головное отделение',
                '-'
            ],
        },
        'Юго-Западный банк (ЮЗБ)': {
            '[Аппарат ЮЗБ]': [
                'Дагестанское ГОСБ №8590',
                'Ингушское ГОСБ №8633',
                'Кабардино-Балкарское ГОСБ №8631',
                'Калмыцкое ГОСБ №8579',
                'Карачаево-Черкесское ГОСБ №8585',
                'Чеченское ГОСБ №8643',
                'Северо-Осетинское ГОСБ №8632',
                '-'
            ]
        },
        'Дальневосточный банк (ДВБ)': {
            '[Головное отделение по Хабаровскому краю ДВБ]': [
                'Биробиджанский ГОСБ №4157',
                'Чукотское головное отделение  ',
                'Головное отделение по Хабаровскому краю',
                'Камчатский ГОСБ №8556',
            ]
        }
    }

    def __init__(
        self,
        df,
        exceptions,
        kpkuo_ker_date_start,
        kpkuo_ker_date_finish,
        keppl_date_start,
        keppl_date_finish,
        kevu_date_start,
        kevu_date_finish,
        kmpk_date_start,
        kmpk_date_finish,
    ):
        self.df = df
        self.exceptions = exceptions
        self.kpkuo_ker_date_start = kpkuo_ker_date_start
        self.kpkuo_ker_date_finish = kpkuo_ker_date_finish
        self.keppl_date_start = keppl_date_start
        self.keppl_date_finish = keppl_date_finish
        self.kevu_date_start = kevu_date_start
        self.kevu_date_finish = kevu_date_finish
        self.kmpk_date_start = kmpk_date_start
        self.kmpk_date_finish = kmpk_date_finish
        self.TB = df.loc[0, 'ТБ/ЦА']
        self.GOSB = df['Подразделение'].unique().tolist()
        self.report = pd.DataFrame(
            columns=[
                'Кпкуо', 'Кэр', 'Кэву', 'Кэппл', 'Кмпк',
                'Кол-во исключений',
                'Кпкуо/Кэр: Кол-во ВУД', 'Кпкуо/Кэр: Передано в суд',
                'Кпкуо/Кэр: Установлено лиц',
                'Кэву: Сумма прич.', 'Кэву: Сумма возм.',
                'Кэппл: Кол-во ВУД', 'Кэппл: Передано в суд(>365)',
                'Кэппл: Передано в суд(<=365)',
                'Кмпк: Кол-во заяв.', 'Кмпк: Кол-во ВУД',
            ],
            index=[self.TB] + list(self.GROUP_GOSB[self.TB].keys()) + self.GOSB,
            data=0,
        )
        self.update()

    def update(self):
        for i, row in self.df.iterrows():
            groups = [self.TB, row['Подразделение']] + [group for group, gosbs in self.GROUP_GOSB[self.TB].items() if row['Подразделение'].strip() in gosbs]
            if (
                row['Статус КЗОП'] == 'Архив' or
                row['Вид события'] == 'Мошенничество в корпоративном кредитовании'
            ):
                continue
            elif row['Номер карточки'] in self.exceptions:
                for group in groups:
                    self.report.loc[group, "Кол-во исключений"] += 1
                continue
            self.check_kmpk(row, groups)
            self.check_keppl(row, groups)
            self.check_kpkuo_ker(row, groups)
            self.check_kevu(row, groups)
        self.calc()

    def check_kmpk(self, row: pd.Series, groups: List = []):
        if row['Потерпевший СБЕР'] == 'false':
            if (
                row['Дата подачи ЗОП'] != '-' and
                row['Дата подачи ЗОП'].date() >= self.kmpk_date_start and
                row['Дата подачи ЗОП'].date() <= self.kmpk_date_finish
            ):
                for group in groups:
                    self.report.loc[group, 'Кмпк: Кол-во заяв.'] += 1
                if row['Дата возбуждения УД'] != '-':
                    for group in groups:
                        self.report.loc[group, 'Кмпк: Кол-во ВУД'] += 1

    def check_kpkuo_ker(self, row: pd.Series, groups: List = []):
        if row['Потерпевший СБЕР'] == 'true':
            if (
                row['Дата возбуждения УД'] != '-' and
                row['Дата возбуждения УД'].date() >= self.kpkuo_ker_date_start and
                row['Дата возбуждения УД'].date() <= self.kpkuo_ker_date_finish
            ):
                for group in groups:
                    self.report.loc[self.TB, 'Кпкуо/Кэр: Кол-во ВУД'] += 1
                if (
                    row['Дата передачи дела в суд первой инстанции'] != 'Отсутствует' and
                    row['Дата передачи дела в суд первой инстанции'].date() <= row['Дата возбуждения УД'].date() + timedelta(days=365)
                ):
                    for group in groups:
                        self.report.loc[group, 'Кпкуо/Кэр: Передано в суд'] += 1
                if row['Подозреваемые']:
                    for group in groups:
                        self.report.loc[group, 'Кпкуо/Кэр: Установлено лиц'] += 1

    def check_keppl(self, row: pd.Series, groups: List = []):
        if row['Потерпевший СБЕР'] == 'true':
            if (
                row['Дата возбуждения УД'] != '-' and
                row['Дата возбуждения УД'].date() >= self.keppl_date_start and
                row['Дата возбуждения УД'].date() <= self.keppl_date_finish
            ):
                for group in groups:
                    self.report.loc[group, 'Кэппл: Кол-во ВУД'] += 1
                if row['Дата передачи дела в суд первой инстанции'] != 'Отсутствует':
                    if row['Дата передачи дела в суд первой инстанции'].date() <= row['Дата возбуждения УД'].date() + timedelta(days=365):
                        for group in groups:
                            self.report.loc[
                                group,
                                'Кэппл: Передано в суд(<=365)'
                            ] += 1
                    else:
                        for group in groups:
                            self.report.loc[
                                group,
                                'Кэппл: Передано в суд(>365)'
                            ] += 1

    def check_kevu(self, row: pd.Series, groups: List = []):
        if row['Потерпевший СБЕР'] == 'true':
            if (
                row['Дата подачи ЗОП'] != '-' and
                row['Дата подачи ЗОП'].date() >= self.kevu_date_start and
                row['Дата подачи ЗОП'].date() <= self.kevu_date_finish
            ):
                if row['Сумма ущерба'] != '-':
                    for group in groups:
                        self.report.loc[
                            group,
                            'Кэву: Сумма прич.'
                        ] += float(row['Сумма ущерба'])
                if row['Ущерб возмещенный'] != '-':
                    for group in groups:
                        self.report.loc[
                            group,
                            'Кэву: Сумма возм.'
                        ] += float(row['Ущерб возмещенный'])

    def calc(self):
        self.report['Кпкуо'] = round(
            100 * self.report['Кпкуо/Кэр: Передано в суд'] / self.report['Кпкуо/Кэр: Кол-во ВУД'], 2
        )
        self.report['Кэр'] = round(
            100 * self.report['Кпкуо/Кэр: Установлено лиц'] / self.report['Кпкуо/Кэр: Кол-во ВУД'], 2
        )
        self.report['Кэппл'] = round(
            100 * self.report['Кэппл: Передано в суд(>365)'] / (self.report['Кэппл: Кол-во ВУД'] - self.report['Кэппл: Передано в суд(<=365)']), 2
        )
        self.report['Кэву'] = round(
            100 * self.report['Кэву: Сумма возм.'] / self.report['Кэву: Сумма прич.'], 2
        )
        self.report['Кмпк'] = round(
            100 * self.report['Кмпк: Кол-во ВУД'] / self.report['Кмпк: Кол-во заяв.'], 2
        )
        self.report.replace([np.inf, -np.inf, np.nan], 0, inplace=True)
        self.report.rename(
            index={'-': 'Аппарат'},
            inplace=True,
            errors='ignore'
        )


class ES_legal():
    GROUP_GOSB = {
        'Среднерусский банк (СРБ)': {
            '[Аппарат СРБ]': [
                'Южное Головное отделение',
                'Восточное Головное отделение',
                'Северное Головное отделение',
                'Западное Головное отделение',
                '-'
            ],
        },
        'Юго-Западный банк (ЮЗБ)': {
            '[Аппарат ЮЗБ]': [
                'Дагестанское ГОСБ №8590',
                'Ингушское ГОСБ №8633',
                'Кабардино-Балкарское ГОСБ №8631',
                'Калмыцкое ГОСБ №8579',
                'Карачаево-Черкесское ГОСБ №8585',
                'Чеченское ГОСБ №8643',
                'Северо-Осетинское ГОСБ №8632',
                '-'
            ]
        },
        'Дальневосточный банк (ДВБ)': {
            '[Головное отделение по Хабаровскому краю ДВБ]': [
                'Биробиджанский ГОСБ №4157',
                'Чукотское головное отделение  ',
                'Головное отделение по Хабаровскому краю',
                'Камчатский ГОСБ №8556',
            ]
        }
    }

    def __init__(
        self,
        df,
        koup_kvu_date_start,
        koup_kvu_date_finish,
        registries
    ):
        self.df = df
        self.koup_kvu_date_start = koup_kvu_date_start
        self.koup_kvu_date_finish = koup_kvu_date_finish
        self.registries = registries
        self.TB = df.loc[0, 'ТБ/ЦА']
        self.GOSB = df['Подразделение'].unique().tolist()
        self.report = pd.DataFrame(
            columns=[
                'Коуп', 'Кву',
                'Коуп: Кол-во', 'Коуп: Кол-во ВУД', 'Коуп: Кол-во искл.', 'Коуп: Кол-во доб.',
                'Кву: Сумма', 'Кву: Кол-во', 'Кву: Кол-во искл.', 'Кву: Кол-во доб.',
            ],
            index=[self.TB] + list(self.GROUP_GOSB[self.TB].keys()) + self.GOSB,
            data=0,
        )
        self.report['Кву: Ущерб возмещенный/Сумма ущерба'] = [[] for _ in self.report.index.values]
        self.update()

    def update(self):
        for i, row in self.df.iterrows():
            groups = [self.TB, row['Подразделение']] + [group for group, gosbs in self.GROUP_GOSB[self.TB].items() if row['Подразделение'].strip() in gosbs]
            if (
                row['Статус КЗОП'] == 'Архив' or
                row['Вид события'] != 'Мошенничество в корпоративном кредитовании' or
                row['Дата подачи ЗОП'] == '-' or
                row['Дата подачи ЗОП'].date() < self.koup_kvu_date_start or
                row['Дата подачи ЗОП'].date() > self.koup_kvu_date_finish
            ):
                continue
            self.check_koup(row, groups)
            self.check_kvu(row, groups)
        self.check_kvu_kvupp_kpupp()
        self.check_koup_zpp_vpp()
        self.calc()

    def check_koup(self, row: pd.Series, groups: List = []):
        for group in groups:
            self.report.loc[group, 'Коуп: Кол-во'] += 1
            if row['Дата возбуждения УД'] != '-' and row['Номер карточки'] not in self.registries['legal_koup_kur']['КОУП_Номер карточки'].values:
                self.report.loc[group, 'Коуп: Кол-во ВУД'] += 1
            if row['Номер карточки'] in self.registries['legal_koup_kur']['КОУП_Номер карточки'].values:
                self.report.loc[group, 'Коуп: Кол-во искл.'] += 1

    def check_kvu(self, row: pd.Series, groups: List = []):
        for group in groups:
            if (
                row['Сумма ущерба'] != '-' and
                row['Номер карточки'] not in self.registries['legal_kvu_kur']['КВУ_Номер карточки'].values
            ):
                if row['Ущерб возмещенный'] == '-':
                    row['Ущерб возмещенный'] = 0
                self.report.loc[group, 'Кву: Сумма'] += float(row['Ущерб возмещенный'])/float(row['Сумма ущерба'])
                self.report.loc[group, 'Кву: Кол-во'] += 1
            if row['Номер карточки'] in self.registries['legal_kvu_kur']['КВУ_Номер карточки'].values:
                self.report.loc[group, 'Кву: Кол-во искл.'] += 1

    def check_kvu_kvupp_kpupp(self):
        for i, row in self.registries['legal_kvu_kvupp_kpupp'].iterrows():
            if row['КВУ_ТБ'] == self.TB:
                groups = [self.TB, row['Подразделение']] + [group for group, gosbs in self.GROUP_GOSB[self.TB].items() if row['Подразделение'].strip() in gosbs]
                if row['КВУ_Номер карточки'] in self.df['Номер карточки']:
                    continue
                for group in groups:
                    self.report.loc[group, 'Кву: Сумма'] += row['КВУПП_Ущерб возмещеннный']/row['КПУПП_Ущерб причиненный']
                    self.report.loc[group, 'Кву: Кол-во'] += 1
                    self.report.loc[group, 'Кву: Кол-во доб.'] += 1

    def check_koup_zpp_vpp(self):
        for i, row in self.registries['legal_koup_zpp_vpp'].iterrows():
            if row['КОУП_ТБ'] == self.TB:
                groups = [self.TB, row['Подразделение']] + [group for group, gosbs in self.GROUP_GOSB[self.TB].items() if row['Подразделение'].strip() in gosbs]
                if row['КОУП_Номер карточки'] in self.df['Номер карточки']:
                    continue
                for group in groups:
                    self.report.loc[group, 'Коуп: Кол-во доб.'] += 1

    def calc(self):
        self.report['Коуп'] = round(
            100 * (self.report['Коуп: Кол-во ВУД'] + self.report['Коуп: Кол-во доб.']) / (self.report['Коуп: Кол-во'] - self.report['Коуп: Кол-во искл.'] + self.report['Коуп: Кол-во доб.']), 2
        )
        self.report['Кву'] = round(
            100 * self.report['Кву: Сумма'] / self.report['Кву: Кол-во'], 2
        )
        self.report.replace([np.inf, -np.inf, np.nan], 0, inplace=True)
        self.report.rename(
            index={'-': 'Аппарат'},
            inplace=True,
            errors='ignore'
        )


if __name__ == '__main__':
    Window()
