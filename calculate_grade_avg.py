import json
import pandas as pd
from collections import defaultdict
import logging
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


def calculate_grade_cut(file_composite, file_strings, subblock):
    # Настройка логгирования
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[logging.StreamHandler()]
    )

    # Устанавливаем допустимое расстояние до полигона (в метрах)
    TOLERANCE_TO_BORDER = 0.2  # например, 0.5 м

    def point_to_segment_distance(px, py, x1, y1, x2, y2):
        """Расстояние от точки (px, py) до отрезка (x1, y1)-(x2, y2)"""
        dx = x2 - x1
        dy = y2 - y1
        if dx == 0 and dy == 0:
            return ((px - x1) ** 2 + (py - y1) ** 2) ** 0.5
        t = max(0, min(1, ((px - x1) * dx + (py - y1) * dy) / (dx ** 2 + dy ** 2)))
        proj_x = x1 + t * dx
        proj_y = y1 + t * dy
        return ((px - proj_x) ** 2 + (py - proj_y) ** 2) ** 0.5

    def is_point_in_polygon(x, y, polygon):
        """Определение, находится ли точка внутри полигона (алгоритм луча)"""
        inside = False
        n = len(polygon)
        j = n - 1
        for i in range(n):
            xi, yi = polygon[i]
            xj, yj = polygon[j]
            if ((yi > y) != (yj > y)):
                slope = (xj - xi) / ((yj - yi) + 1e-12)
                x_intersect = slope * (y - yi) + xi
                if x < x_intersect:
                    inside = not inside
            j = i
        return inside

    def is_point_in_polygon_with_tol(x, y, polygon, tol=TOLERANCE_TO_BORDER):
        """Проверка: внутри полигона или рядом с границей"""
        if is_point_in_polygon(x, y, polygon):
            return True
        n = len(polygon)
        for i in range(n):
            x1, y1 = polygon[i]
            x2, y2 = polygon[(i + 1) % n]
            dist = point_to_segment_distance(x, y, x1, y1, x2, y2)
            if dist <= tol:
                return True
        return False


    def notification_to_close_file(df, name_file, index=False):
        logging.debug(f" === Создание файла Excel: {name_file}.xlsx === ")
        try:
            df.to_excel(f"{name_file}.xlsx", index=index)
            logging.debug(f" === Файл Excel {name_file}.xlsx --> Создан. === ")
        except PermissionError:
            logging.error(f" === Не удалось создать файл Excel, необходимо закрыть файл: {name_file}.xlsx === ")

    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)
    pd.set_option('display.max_colwidth', None)


    # DATA_NAME_FILE_ASSAY = "data_assay.json" # Файл данных о скважинах|композиты|опробование

    # DATA_NAME_FILE_STRINGS = "data_strings.json" # Файл данных о полигонах/стрингах

    EXPLORATION_BLOCK  = subblock

    # # === ЭТАП 1: Загрузка данных ===

    # # 1.1 == Загрузка файла скважин|композиты|опробование: ==
    # try:
    #     logging.info(f"{"="*50}Этап 1: Загрузка данных...{"="*50}")
    #     logging.info(f"Подэтап 1.1: Загрузка JSON-файла {DATA_NAME_FILE_ASSAY}")
    #     with open(DATA_NAME_FILE_ASSAY , "r", encoding="utf-8") as f:
    #         data_assay = json.load(f)
    #     if data_assay:
    #         logging.info(f"JSON-файл {DATA_NAME_FILE_ASSAY} успешно загружен!")
    #     else:
    #         logging.warning(f"JSON файл {DATA_NAME_FILE_ASSAY} пуст (пустой список или словарь).")
    # except json.JSONDecodeError:
    #     logging.critical(f"JSON-файл {DATA_NAME_FILE_ASSAY} повреждён или полностью пуст.")
    # except FileNotFoundError:
    #     logging.critical(f"Файл {DATA_NAME_FILE_ASSAY} не найден.")

    # # 1.2 == Загрузка файла стринги|полигоны в которые входят скважины: ==
    # try:
    #     logging.info(f"Подэтап 1.2: Загрузка JSON-файла {DATA_NAME_FILE_STRINGS}")
    #     with open(DATA_NAME_FILE_STRINGS , "r", encoding="utf-8") as f:
    #         data_strings = json.load(f)
    #     if data_strings:
    #         logging.info(f"JSON-файл {DATA_NAME_FILE_STRINGS} успешно загружен!")
    #     else:
    #         logging.warning(f"JSON файл {DATA_NAME_FILE_STRINGS} пуст (пустой список или словарь).")
    # except json.JSONDecodeError:
    #     logging.critical(f"JSON-файл {DATA_NAME_FILE_STRINGS} повреждён или полностью пуст.")
    # except FileNotFoundError:
    #     logging.critical(f"Файл {DATA_NAME_FILE_STRINGS} не найден.")

    # === ЭТАП 2: Определение вхождение и запись скважин в полигоны ===


    # Перевод словарей в DataFrame
    data_assay = file_composite
    data_strings = file_strings
    # 2.1 == Создание номера и запись субблока: ==
    try:

        logging.info(f"{"="*50}Этап 2: Создание номера и запись субблока запущено...{"="*50}")
        logging.info(f"Подэтап 2.1:Создание номера и запись субблока")

        if "JOIN" not in data_strings.columns:
            logging.critical("Столбец 'JOIN' не найден в data_strings")

        else:
            data_strings["JOIN"] = data_strings["JOIN"].astype(str)
            string = EXPLORATION_BLOCK + "_" + data_strings["JOIN"]
            data_strings["Субблок"] = string
            logging.debug(f"Номера субблоков {set(string)} сформированы!")
            logging.info(f"Субблока успешно созданы и записаны!")
            
    except Exception as e:
        logging.critical(f"Ошибка при создании и записи субблоков: {e}")

    # 2.2 == Подготовка данных о полигонах/Создание словаря с координатами и номерами субблоков: ==
    try:

        polygons = defaultdict(list)
        logging.info(f"Подэтап 2.2: Подготовка данных координат полигонов")
        logging.info(f"Формирование координат полигонов.")

        for i, row in data_strings.iterrows():
            number_block = row["Субблок"]
            north = pd.to_numeric(row["NORTH"], errors="coerce")
            east = pd.to_numeric(row["EAST"], errors="coerce")
            polygons[number_block].append((east, north))  # (X, Y)
            logging.debug(f"Субблок - {number_block} - координаты {(east, north)} добавлены")

        logging.info(f"Координаты успешно сформированы!.")
    except Exception as e:
        logging.critical(f"Не получилось сформировать координаты полигонов: {e}")


    # 2.3  == Определение вхождение скважины в полигон: ==
    logging.info(f"Подэтап 2.3: Определение вхождение скважины в полигон")
    for i, row_assay in data_assay.iterrows():
        x = pd.to_numeric(row_assay["X"], errors="coerce") # Север
        y = pd.to_numeric(row_assay["Y"], errors="coerce") # Восток
        hole = row_assay["№ пробы полевой"]

        if pd.isna(row_assay["X"]) or pd.isna(row_assay["Y"]):
            logging.warning("Обнаружены некорректные координаты X или Y (NaN) в строке.")

        found = False
        if row_assay["ЭБ"] == EXPLORATION_BLOCK:
            for number_block, coords in polygons.items():

                if is_point_in_polygon_with_tol(x, y, coords):
                    data_assay.at[i, "Субблок"] = number_block
                    found = True
                    if found:
                        logging.debug(f"Скважина :{hole[:-2]} входит в субблок ==> {number_block}")
                    break

            if not found:
                data_assay.at[i,"Субблок"] = 0
                logging.warning(f"Скважина :{hole[:-2]} никуда не вошла")

    logging.info(f"Закончен процесс присвоение скважинам субблока!")

    # === ЭТАП 3: Определение метода урагана ===

    logging.info(f"{"="*50}ЭТАП 3: Определение метода урагана{"="*50}")

    filtered_eb = data_assay[data_assay["ЭБ"] == EXPLORATION_BLOCK]

    # 3.1  == Подсчёт скважин входящие по каждому субблоку и подготовка данных к вычислениям == 
    logging.info(f"Подэтап 3.1: Подсчёт количества скважин в каждом субблоке")

    quantity_hole = filtered_eb["Субблок"].value_counts()

    logging.info(f"Подсчёт закончен!")

    df_counts = quantity_hole.reset_index()
    df_counts.columns = ["Субблок", "Количество"]

    # Фильтруем блока, в которых меньше или равно 7
    small_blocks = quantity_hole[quantity_hole <= 7]

    # Получаем список субблоков с количеством <= 7
    small_blocks_list = small_blocks.index.tolist()

    # Фильтруем строки из основной таблицы
    small_blocks_rows = data_assay[data_assay["Субблок"].isin(small_blocks_list)].copy()

    # Заменяем пустые строки на NaN, чтобы не вылезала ошибка, а также если числовые значение записаны как str
    small_blocks_rows["S_AU_2.5"] = None
    small_blocks_rows["Au, г/т"] = pd.to_numeric(small_blocks_rows["Au, г/т"], errors="coerce")
    small_blocks_rows["S_AU_2.5"] = pd.to_numeric(small_blocks_rows["Au, г/т"], errors="coerce")

    if small_blocks_rows["Au, г/т"].isna().any():
        logging.warning("Обнаружены некорректные значения в столбце 'Au, г/т' после преобразования в float")

    # Урезаем все содержание больше 2.5
    small_blocks_rows.loc[small_blocks_rows["Au, г/т"] > 2.5, "S_AU_2.5"] = 2.5

    # Перенос данных в столбец "Урезка1"
    small_blocks_rows["Урезка1"] = small_blocks_rows["S_AU_2.5"]

    # 3.2  == Метод 1-й, если скважин <= 7 то урезаем содержание > 1.5 до 1.5 == 
    logging.info(f"Подэтап 3.2: Урезка содержаний больше 1.5(1-й метод)")

    small_blocks_rows.loc[small_blocks_rows["Урезка1"] > 1.5, "Урезка1"] = 1.5
    logging.info(f"Урезка1 завершена!")


    # 3.3 == Вычисление среднего после урезки1 ==
    # вычисление среднего по столбцу Урезка1 после замены содержаний выше 1.5 на 1.5
    logging.info(f"Подэтап 3.3: Вычисление среднего содержание по 'Урезка1'")
    mean_au_by_block = small_blocks_rows.groupby("Субблок")["Урезка1"].mean()

    # Переобразование из serias в dataframe
    df_average_au = mean_au_by_block.reset_index()
    df_average_au.columns = ["Субблок", "Среднее1_по_Урезка1"]

    for i, rows in df_average_au.iterrows():
        logging.debug(f"Среднее содежание по 'Урезка1' субблок: {rows["Субблок"]} = {rows["Среднее1_по_Урезка1"]:2f}!")

    # Добавление столбцов
    small_blocks_rows["Урезка2"] = None
    small_blocks_rows["Урезка3"] = None
    df_average_au["Среднее2_по_Урезка2"] = None
    df_average_au["Среднее3_по_Урезка3"] = None

    # копирование данных в столбце AU_CUT
    small_blocks_rows["AU_CUT"] = small_blocks_rows["Au, г/т"]

    for i, row in df_average_au.iterrows():
        block = row["Субблок"]
        avg_au = row["Среднее1_по_Урезка1"]
        block_mask = small_blocks_rows["Субблок"] == block

        if avg_au > 0.3:

            # 3.4 == Урезка2 ==
            # Урезка2 по Среднее1
            logging.info(f"Подэтап 3.4: Вычисление среднего содержание по 'Урезка2")
            small_blocks_rows.loc[block_mask, "Урезка2"] = small_blocks_rows.loc[block_mask, "Урезка1"].clip(upper=avg_au)
            mean_cut2 = small_blocks_rows.loc[block_mask, "Урезка2"].astype(float).mean()
            df_average_au.at[i, "Среднее2_по_Урезка2"] = mean_cut2
            logging.debug(f"Среднее содержание по 'Урезка2' субблок: {block} = {mean_cut2:2f}!")

            # 3.5 == Урезка3 ==
            logging.info(f"Подэтап 3.5: Вычисление среднего содержание по 'Урезка3")

            # Урезка3 по Среднее2
            small_blocks_rows.loc[block_mask, "Урезка3"] = small_blocks_rows.loc[block_mask, "Урезка2"].clip(upper=mean_cut2)
            mean_cut3 = small_blocks_rows.loc[block_mask, "Урезка3"].astype(float).mean()
            df_average_au.at[i, "Среднее3_по_Урезка3"] = mean_cut3

            # По каждой скважине добавляем значение Урезка3 в AU_CUT
            for j, r in small_blocks_rows.loc[block_mask].iterrows():
                au_cut_value = r["Урезка3"]
                small_blocks_rows.at[j, "AU_CUT"] = au_cut_value
                logging.debug(f"Скважина: {r['Скважина2']} | AU_CUT = {au_cut_value:.3f}")

            logging.debug(f"Среднее содержание по 'Урезка3' субблок: {block} = {mean_cut3:2f}!")

            # ==Вычисление итоговых содержаний ==
            logging.info(f"Вычисление итогового содержания")

            # 20-й перцентиль по блоку
            p20 = small_blocks_rows.loc[block_mask, "Урезка3"].astype(float).quantile(0.2)

            # Итоговое значение
            if avg_au > 1:
                finally_avg_au = p20 * 0.7
            elif p20 < 0.3:
                finally_avg_au = 0.3
            else:
                finally_avg_au = p20

            df_average_au.at[i, "Итоговое"] = finally_avg_au
            logging.debug(f"Итоговое содержание: {block} = {finally_avg_au:2f}!")

        else:
            small_blocks_rows.loc[block_mask, "Урезка2"] = None
            small_blocks_rows.loc[block_mask, "Урезка3"] = None
            # Если ни одна из урезок не была рассчитана — сохраняем avg_au как итоговое
            if pd.isna(df_average_au.at[i, "Среднее2_по_Урезка2"]) and pd.isna(df_average_au.at[i, "Среднее3_по_Урезка3"]):
                df_average_au.at[i, "Итоговое"] = avg_au

    logging.info(f"Урезка2 завершена!")
    logging.info(f"Урезка3 завершена!")
    logging.info(f"Итоговые значение завершены!")

    # === ЭТАП 4: Урезка ураганных значений в блоках больше 7 скважин ===
    logging.info(f"{"="*50}ЭТАП 4: Урезка ураганных значений в блоках больше 7 скважин{"="*50}")

    # Подэтап 4.1: Вычисление метрограмма
    logging.info(f"Подэтап 4.1: Вычисление метрограмма")

    # Фильтруем блока, в которых больше 7 скважин
    big_blocks = quantity_hole[quantity_hole > 7]

    # Создаем список с субблоками в которых входят больше 7 скважин
    big_blocks_list = big_blocks.index.tolist()

    # Фильтрация данных по столбцу "Субблок"
    big_blocks_rows = data_assay[data_assay["Субблок"].isin(big_blocks_list)].copy()

    big_blocks_rows["S_AU_2.5"] = None

    big_blocks_rows["Au, г/т"] = pd.to_numeric(big_blocks_rows["Au, г/т"], errors="coerce")
    big_blocks_rows["S_AU_2.5"] = pd.to_numeric(big_blocks_rows["Au, г/т"], errors="coerce")

    # Урезаем все содержание больше 2.5
    big_blocks_rows.loc[big_blocks_rows["S_AU_2.5"] > 2.5, "S_AU_2.5"] = 2.5

    sum_metrogram_and_sample = {}

    sum_metrogram_and_sample = pd.DataFrame.from_dict(sum_metrogram_and_sample, orient="index", columns=["Сумма_метрограмма", "Сумма_длин"])
    sum_metrogram_and_sample.reset_index(inplace=True)
    sum_metrogram_and_sample.columns = ["Субблок", "Сумма_метрограмма", "Сумма_длин"]


    # Вычисление и импорт метрограмма
    for i, row in big_blocks_rows.iterrows():
        grade_au = row["S_AU_2.5"]
        sample_long = row["Длина"]
        block = row["Субблок"]
        hole = row["Скважина2"]
        metrogram = float(grade_au) * float(sample_long)
        big_blocks_rows.at[i, "Метрограмм"] = metrogram
        logging.debug(f"Субблок: {block}| Скважина:{hole}| Метрограмм:{metrogram}")

    big_blocks_rows["Метрограмм"] = pd.to_numeric(big_blocks_rows["Метрограмм"], errors="coerce")
    big_blocks_rows["Длина"] = pd.to_numeric(big_blocks_rows["Длина"], errors="coerce")

    logging.info(f"Завершение вычисление метрограмма!")

    # Подэтап 4.2: Вычисление суммы метрограмма и длин проб блока
    logging.info(f"Подэтап 4.2: Вычисление суммы метрограмма и длин проб блока")

    # Группируем и считаем суммы
    sum_metrogram_and_sample = (
        big_blocks_rows
        .groupby("Субблок")[["Метрограмм", "Длина"]]
        .sum()
        .reset_index()
    )

    sum_metrogram_and_sample.columns = ["Субблок", "Сумма_метрограмма", "Сумма_длин"]

    for i, row in sum_metrogram_and_sample.iterrows():
        logging.debug(f"{row["Субблок"]}: Сумма_метрограмма= {row["Сумма_метрограмма"]}| Сумма_длин_проб = {row["Сумма_длин"]}")
    logging.info(f"Суммы подсчитаны!")

    sum_metrogram_and_sample["Ураган_метрограмм"] = None

    # Подэтап 4.3: Вычисление 10% метрограмма
    logging.info(f"Подэтап 4.3: Вычисление 10% метрограмма")
    for i, row in sum_metrogram_and_sample.iterrows():
        sum_metrogram = row["Сумма_метрограмма"]
        sum_metrogram_and_sample.at[i,"Ураган_метрограмм"] = sum_metrogram * 0.05
        logging.debug(f"Метрограмм 10%: субблок = {row["Субблок"]} | метрограмм 10% = {sum_metrogram * 0.1}")
    logging.info("10% метрограмм подсчитан!")

    big_blocks_rows["Метрограмм"] = pd.to_numeric(big_blocks_rows["Метрограмм"], errors="coerce")

    # Копируем исходные значения в новый столбец
    big_blocks_rows["Урезанный_метрограмм"] = big_blocks_rows["Метрограмм"]

    # Подэтап 4.4: Вычисление метрограмма выше 10% и урезка
    logging.info(f"Подэтап 4.4: Вычисление метрограмма выше 10% и урезка")
    # Применяем обрезку по каждому блоку

    for i, row in sum_metrogram_and_sample.iterrows():
        block = row["Субблок"]
        hurricane_limit = row["Ураган_метрограмм"]

        # Маска для строк этого блока, где нужно урезать
        cut_mask = (big_blocks_rows["Субблок"] == block) & (big_blocks_rows["Метрограмм"] > hurricane_limit)

        # Урезаем метрограмм
        big_blocks_rows.loc[cut_mask, "Урезанный_метрограмм"] = hurricane_limit

        for _, r in big_blocks_rows.loc[cut_mask].iterrows():
            logging.debug(
                f"Урезан метрограмм: субблок = {
                    block} | скважина = {r['Скважина2']} | исходный = {r['Метрограмм']:.2f} | урезанный = {hurricane_limit:.2f}"
            )


    # === ЭТАП 5: Урезка скв по 4 скв > 1г/т и по 10% для скв < 1г/т  ===

    logging.info(f"{"="*50}ЭТАП 5: Урезка скв по 4 скв > 1г/т и по 10% для скв < 1г/т{"="*50}")

    # Подэтап 5.1: Фильтрование данных выше 10% урагана
    logging.info("Подэтап 5.1: Фильтрование данных: выше 10% урагана")
    filtered_metrogram = big_blocks_rows[
        big_blocks_rows["Метрограмм"] > big_blocks_rows["Урезанный_метрограмм"]
    ]

    for i, row in filtered_metrogram.iterrows():
        logging.debug(f"Потенциальная скважина для урезки (по 4-м скважинам): субблок = {
            row['Субблок']} | скважина = {row["Скважина2"]} | {row['Au, г/т']}")

    # Создаем отдельный словарь с нумерацией скважин и координатами по блоку
    hole_coords = {
        row["Скважина2"]: (float(row["X"]), float(row["Y"]))
        for _, row in filtered_eb.iterrows()
    }

    logging.info(f"Фильтрование данных завершено!")

    # Подэтап 5.2: Поиск ближайших 4-х скважин
    logging.info(f"Подэтап 5.2: Поиск ближайших 4-х скважин")

    # Радиус поиска ближайших скважин
    radius = 10.2
    logging.info(f"Радиус поиска: {radius} метров.")

    nearby_sorted = {}
    # Для каждой строки из filtered_metrogram — ищем ближайшие скважины из всего набора
    for _, row in filtered_metrogram.iterrows():
        current_hole = row["Скважина2"]
        x0 = float(row["X"])
        y0 = float(row["Y"])
        # grade_au = float(row["Au, г/т"])

        # Находим ближайшие, исключая саму скважину
        nearby = []
        for other_hole, (x, y) in hole_coords.items():
            if other_hole == current_hole:
                continue
            dx = x0 - x
            dy = y0 - y
            distance = (dx**2 + dy**2)**0.5
            if distance <= radius:
                nearby.append(other_hole)

        # Сортируем по расстоянию
        nearby_sorted[current_hole] = nearby[:4]

    rows_nearby = []
    for main_hole, neighbors in nearby_sorted.items():
        for idx, neighbor in enumerate(neighbors, start=1):
            rows_nearby.append({
                "Скважина": main_hole,
                "Ближайшая_скважина": neighbor,
                "Номер_по_порядку": idx
            })

    df_nearby = pd.DataFrame(rows_nearby)

    near_hole_dict = defaultdict(list)
    for i, row in df_nearby.iterrows():
        hole = row["Скважина"]
        near_hole = row["Ближайшая_скважина"]
        near_hole_dict[hole].append(near_hole)

    for hole, near_holes in near_hole_dict.items():
        logging.debug(f"Скважина = {hole} => ближайшие 4-е скважины: {near_holes}")

    logging.info(f"Поиск скважин закончен!")

    # Подэтап 5.3: Фильтрование данных для урезки по 4 скважинам и определение метода
    logging.info(f"Подэтап 5.3: Фильтрование данных для урезки по 4 скважинам")

    big_blocks_rows["Скважина2"] = pd.to_numeric(big_blocks_rows["Скважина2"], errors="coerce")
    big_blocks_rows["S_AU_2.5"] = pd.to_numeric(big_blocks_rows["S_AU_2.5"], errors="coerce")

    big_blocks_rows["Метод_урезки"] = None

    # Выбор метода урезки ураганных значение, если содержание Au скважины имеет > 1 г/т
    # и имеет ближайших скважин >= 3, то присваивается метод №1(урезка по 4-м скважинам),
    # в ином случае метод №2(10% метод)
    big_blocks_rows["AU_CUT"] = big_blocks_rows["S_AU_2.5"]

    for current_hole, near_holes in near_hole_dict.items():
        mask = big_blocks_rows["Скважина2"] == int(current_hole)
        matched_rows = big_blocks_rows[mask]

        if (matched_rows["S_AU_2.5"] >= 1).any() and (len(near_holes) >= 3):
            big_blocks_rows.loc[mask, "Метод_урезки"] = 1
            logging.debug(f"{current_hole} метод урезки ==> 1")
        elif (matched_rows["S_AU_2.5"] < 1).any() or (len(near_holes) < 3) :
            big_blocks_rows.loc[mask, "Метод_урезки"] = 2
            logging.debug(f"{current_hole} метод урезки ==> 2")
        else:
            big_blocks_rows.loc[mask, "Метод_урезки"] = 0
            logging.warning(f"{current_hole} метод урезки не определён (возможно NaN или нештатные данные)")


    for hole_cut, list_nearby_holes in near_hole_dict.items():
        mask_hole_cut = big_blocks_rows["Скважина2"] == int(hole_cut)
        
        if mask_hole_cut.sum() == 0:
            logging.warning(f"Скважина {hole_cut} не найдена в данных. Пропущена.")
            continue

        method = big_blocks_rows.loc[mask_hole_cut, "Метод_урезки"].values[0]

        if method == 1:
            logging.debug(f"Скважина {hole_cut} | Метод урезки = 1 (по 4 соседним скважинам)")
            au_list = []
            for near in list_nearby_holes:
                mask_near = big_blocks_rows["Скважина2"] == int(near)
                if not big_blocks_rows.loc[mask_near, "AU_CUT"].empty:
                    au_value = big_blocks_rows.loc[mask_near, "AU_CUT"].values[0]
                    au_list.append(au_value)
                    logging.debug(f"Соседняя скв. {near}: содержание Au = {au_value}")
                else:
                    logging.warning(f"Соседняя скв. {near} не найдена. Пропущена.")

            if au_list:
                avg_au = sum(au_list) / len(au_list)
                big_blocks_rows.loc[mask_hole_cut, "AU_CUT"] = avg_au
                logging.info(f"Скважина {hole_cut}: среднее содержание соседей = {avg_au:.3f}")
            else:
                logging.warning(f"Скважина {hole_cut}: не удалось найти ни одного соседа с содержанием Au.")

        elif method == 2:
            logging.debug(f"Скважина {hole_cut} | Метод урезки = 2 (по 10% метрограмму)")
            block_cut = big_blocks_rows.loc[mask_hole_cut, "Субблок"].values[0]
            block_mask = sum_metrogram_and_sample["Субблок"] == block_cut

            if block_mask.any():
                hurricane_metrogram = sum_metrogram_and_sample.loc[block_mask, "Ураган_метрограмм"].values[0]

                # Берем длину текущей скважины
                sample_length = big_blocks_rows.loc[mask_hole_cut, "Длина"].values[0]

                if sample_length and sample_length > 0:
                    au_cut = hurricane_metrogram / sample_length
                    big_blocks_rows.loc[mask_hole_cut, "Урезанный_метрограмм"] = hurricane_metrogram
                    big_blocks_rows.loc[mask_hole_cut, "AU_CUT"] = au_cut
                    logging.info(
                        f"Скважина {hole_cut}: урезка по блоку {block_cut} | "
                        f"метрограмм = {hurricane_metrogram:.2f}, длина пробы = {sample_length:.2f}, "
                        f"AU_CUT = {au_cut:.3f}"
                    )
                else:
                    logging.error(f"Скважина {hole_cut}: длина пробы равна 0 или отсутствует")
            else:
                logging.error(f"Скважина {hole_cut}: субблок {block_cut} не найден в sum_metrogram_and_sample")
    logging.info(f"Урезка ураганнов закончена!")


    notification_to_close_file(small_blocks_rows, "1) урезка_меньше_7", True)
    notification_to_close_file(df_nearby, "2) ближайшие_скважины", False)
    notification_to_close_file(filtered_metrogram, "3) ураганы_метрограмм", True)
    notification_to_close_file(sum_metrogram_and_sample, "4) сумма_метрограмма_и_длин", False)
    notification_to_close_file(big_blocks_rows, "5) урезка_больше_7", True)
    notification_to_close_file(df_average_au, "6) средние_по_субблокам", False)


    # Объединяем данные из блоков <=7 и >7 скважин
    final_df = pd.concat([small_blocks_rows, big_blocks_rows], ignore_index=True)

    # Сохраняем в Excel
    output_filename = "7) объединённые_результаты.xlsx"
    try:
        final_df.to_excel(output_filename, index=False)
        final_df.to_json("объединённые_результаты.json", index=False)
        logging.info(f"Итоговый объединённый файл сохранён: {output_filename}")
    except PermissionError:
        logging.error(f"Файл {output_filename} открыт. Закрой его перед сохранением.")


