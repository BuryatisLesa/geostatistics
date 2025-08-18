import pandas as pd
from geostatistics.geostatistics import GeoStatisctics as gs
import logging
import os
from geostatistics.strings import Strings
from collections import defaultdict

def cutGrade(pathFileAssay, pathFileStrings, EXPLORATION_BLOCK):

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[logging.StreamHandler()]
    )
    def exportFileXlsx(df, name_file, index=False):
        logging.debug(f" === Создание файла Excel: {name_file}.xlsx === ")

        base_dir = os.path.dirname(os.path.abspath(__file__))
        folder_path = os.path.join(base_dir, "Scripts_files")
        os.makedirs(folder_path, exist_ok=True)

        if not name_file.lower().endswith(".xlsx"):
            name_file += ".xlsx"

        file_path = os.path.join(folder_path, name_file)

        try:
            df.to_excel(file_path, index=index)
            logging.debug(f" === Файл Excel {name_file} --> Создан. === ")
        except PermissionError:
            logging.error(f" === Не удалось создать файл Excel. Закройте файл: {name_file} и попробуйте снова. === ")

    fileAssay = pd.read_excel(pathFileAssay)
    fileStrings = pd.read_excel(pathFileStrings)

    # 2.1 == Создание номера и запись субблока: ==
    try:
        logging.info(f"{"="*50}Этап 2: Создание номера и запись субблока запущено...{"="*50}")
        logging.info(f"Подэтап 2.1:Создание номера и запись субблока")

        if "JOIN" not in fileStrings.columns:
            logging.critical("Столбец 'JOIN' не найден в data_strings")

        else:
            fileStrings["JOIN"] = fileStrings["JOIN"].astype(str)
            string = EXPLORATION_BLOCK + "-" + fileStrings["JOIN"]
            fileStrings["Субблок"] = string
            logging.debug(f"Номера субблоков {set(string)} сформированы!")
            logging.info(f"Субблока успешно созданы и записаны!")
            
    except Exception as e:
        logging.critical(f"Ошибка при создании и записи субблоков: {e}")

    exportFileXlsx(fileStrings, "1)" + EXPLORATION_BLOCK + "_strings", index=False)

    for index , Z in enumerate(fileAssay["Z"]):
        for minHorizont in range(650, 1165, 5):
            maxHorizont = minHorizont + 5
            if minHorizont < Z < maxHorizont:
                fileAssay.loc[index, "Горизонт"] = minHorizont
    
    exportFileXlsx(fileAssay, "2)Горизонт", index=False)

    # 2.2 == Подготовка данных о полигонах/Создание словаря с координатами и номерами субблоков: ==
    try:

        polygons = defaultdict(list)
        logging.info(f"Подэтап 2.2: Подготовка данных координат полигонов")
        logging.info(f"Формирование координат полигонов.")

        for i, row in fileStrings.iterrows():
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
    for i, row_assay in fileAssay.iterrows():
        x = pd.to_numeric(row_assay["X"], errors="coerce") # Север
        y = pd.to_numeric(row_assay["Y"], errors="coerce") # Восток
        hole = row_assay["№ пробы полевой"]

        if pd.isna(row_assay["X"]) or pd.isna(row_assay["Y"]):
            logging.warning("Обнаружены некорректные координаты X или Y (NaN) в строке.")

        found = False
        if row_assay["ЭБ"] == EXPLORATION_BLOCK:
            for number_block, coords in polygons.items():

                if Strings.is_point_in_polygon_with_tol(x, y, coords):
                    fileAssay.at[i, "Субблок"] = number_block
                    found = True
                    if found:
                        logging.debug(f"Скважина :{hole[:-2]} входит в субблок ==> {number_block}")
                    break

            if not found:
                fileAssay.at[i,"Субблок"] = 0
                logging.warning(f"Скважина :{hole[:-2]} никуда не вошла")

    logging.info(f"Закончен процесс присвоение скважинам субблока!")

    exportFileXlsx(fileAssay, "3)Композиты_CЭР_5м_присвоение_субблоков_скважинам")
        

    #Выборка проб
    filtredFileAssay = gs.filteredData(file=fileAssay, field="ЭБ", filter=EXPLORATION_BLOCK)
    exportFileXlsx(filtredFileAssay, "4)Выборка_проб", index=False)

    #Сортировка проб по субблокам
    filtredFileAssay = filtredFileAssay.sort_values("Субблок", ascending=True)
    exportFileXlsx(filtredFileAssay, "5)Выборка_проб_отсортированы_по_субблокам", index=False)

    logging.info("Подэтап 3.1: Подсчёт количества скважин в каждом субблоке")

    # Подсчёт количества каждого Субблока
    quantityHoleInStrings = filtredFileAssay["Субблок"].value_counts()
    logging.info("Подсчёт закончен!")

    # Сортировка по имени Субблока
    quantityHoleInStrings = quantityHoleInStrings.sort_index(ascending=True)

    # Превращаем в DataFrame и добавляем столбец 'Субблок' если нужно
    quantityHoleInStrings = quantityHoleInStrings.reset_index()
    quantityHoleInStrings.columns = ["Субблок", "ЧАСТОТА"]

    # Экспорт в Excel
    exportFileXlsx(quantityHoleInStrings, "6)Подсчёт_проб_по_блокам", index=False)
    # Присвоение частоты проб выборке
    filtredFileAssay["ЧАСТОТА"] = filtredFileAssay["Субблок"].map(
        filtredFileAssay["Субблок"].value_counts()
    )
    exportFileXlsx(filtredFileAssay, "7)Выборка_проб_присвоение_частоты", index=False)

    # Вычисление метода урезки и присвоение выборке
    filtredFileAssay.loc[filtredFileAssay["ЧАСТОТА"] <= 7, "Метод урезки"] = 1
    filtredFileAssay.loc[filtredFileAssay["ЧАСТОТА"] > 7, "Метод урезки"] = 2
    exportFileXlsx(filtredFileAssay, "8)Выборка_проб_присвоение_метода_урезки")

    # --- Урезка1 ---
    filtredFileAssay["Урезка1"] = pd.NA
    MASK_METHOD_ONE = filtredFileAssay["Метод урезки"] == 1

    # Копируем данные в столбец урезка1, если содержание золота выше 1.5 г/т, тогда понижаем до 1.5
    filtredFileAssay.loc[MASK_METHOD_ONE, "Урезка1"] = filtredFileAssay.loc[MASK_METHOD_ONE, "Au, г/т"].clip(upper=1.5)
    
    # Вычисляем среднее по каждому Субблок и добавляем в исходный DataFrame
    filtredFileAssay["Среднее1"] = filtredFileAssay.groupby("Субблок")["Урезка1"].transform("mean")

    # Фильтруем строки с заполненным Среднее1
    filteredRowsAverage = filtredFileAssay[filtredFileAssay["Среднее1"].notna()]

    # Берём первую строку каждого Субблок
    firstRowsAverage = filteredRowsAverage.groupby("Субблок").first().reset_index()
    exportFileXlsx(firstRowsAverage, "9)Среднее1", index=False)
    exportFileXlsx(filtredFileAssay, "10)Выборка_проб_среднее1", index=False)

    # Вычисление индекса минимального содержание в субблоках
    filteredMinGradeInBlock = filtredFileAssay[filtredFileAssay["Урезка1"].notna()]
    idxMinGradeInBlock = filteredMinGradeInBlock.groupby("Субблок")["Урезка1"].idxmin()
    minGradeInBlockOne = filtredFileAssay.loc[idxMinGradeInBlock].reset_index(drop=True)
    exportFileXlsx(minGradeInBlockOne, "11)Минимальное1", index=False)

    minGradeInBlockOne.loc[minGradeInBlockOne["Урезка1"] >= 0.4, "Урезка1"] = 0.3

    min_values_map = minGradeInBlockOne.set_index("Скважина2")["Урезка1"].to_dict()

    # Присваиваем только тем строкам, где Скважина2 есть в minGradeInBlockOne
    filtredFileAssay.loc[
        filtredFileAssay["Скважина2"].isin(minGradeInBlockOne["Скважина2"]), "Урезка1"
    ] = filtredFileAssay.loc[
        filtredFileAssay["Скважина2"].isin(minGradeInBlockOne["Скважина2"]), "Скважина2"
    ].map(min_values_map)

    # --- Урезка2 ---

    
    # Копируем данные с Урезка1 и урезаем содержания выше Среднее1
    upperValuesOne = filtredFileAssay.groupby("Субблок")["Среднее1"].transform("first")
    filtredFileAssay.loc[MASK_METHOD_ONE, "Урезка2"] = filtredFileAssay.loc[MASK_METHOD_ONE, "Урезка1"].clip(upper=upperValuesOne.loc[MASK_METHOD_ONE])

    # Вычисляем Среднее2
    filtredFileAssay["Среднее2"] = filtredFileAssay.groupby("Субблок")["Урезка2"].transform("mean")

    # Фильтруем строки с заполненным Среднее2
    filteredRowsAverage = filtredFileAssay[filtredFileAssay["Среднее2"].notna()]

    # Берём первую строку каждого Субблок
    TwoRowsAverage = filteredRowsAverage.groupby("Субблок").first().reset_index()
    exportFileXlsx(TwoRowsAverage, "12)Среднее2", index=False)
    exportFileXlsx(filtredFileAssay, "13)Выборка_проб_среднее2", index=False)

    # --- Урезка3 ---

    # Копируем данные с Урезка2 и урезаем содержания выше Среднее2
    upperValuesTwo = filtredFileAssay.groupby("Субблок")["Среднее2"].transform("first")
    filtredFileAssay.loc[MASK_METHOD_ONE, "Урезка3"] = filtredFileAssay.loc[MASK_METHOD_ONE, "Урезка2"].clip(upper=upperValuesTwo.loc[MASK_METHOD_ONE])

    # Вычисляем Среднее3
    filtredFileAssay["Среднее3"] = filtredFileAssay.groupby("Субблок")["Урезка3"].transform("mean")

    # Фильтруем строки с заполненным Среднее3
    filteredRowsAverage = filtredFileAssay[filtredFileAssay["Среднее3"].notna()]

    # Берём первую строку каждого Субблок
    ThreeRowsAverage = filteredRowsAverage.groupby("Субблок").first().reset_index()
    exportFileXlsx(ThreeRowsAverage, "14)Среднее3", index=False)
    exportFileXlsx(filtredFileAssay, "15)Выборка_проб_среднее3", index=False)

    # Вычисление 20 процентиль
    filtredFileAssay["20 Процентиль"] = filtredFileAssay.groupby("Субблок")["Урезка3"].transform(lambda x: x.quantile(0.2))
    exportFileXlsx(filtredFileAssay, "16)20_Процентиль", index=False)

    mean1 = filtredFileAssay.groupby("Субблок")["Среднее1"].transform("first")
    mean3 = filtredFileAssay.groupby("Субблок")["Среднее3"].transform("first")
    p20 = filtredFileAssay.groupby("Субблок")["20 Процентиль"].transform("first")

    # Итоговое содержание для AU_CUT
    filtredFileAssay.loc[mean1 <= 0.3, "AU_CUT"] = mean1
    filtredFileAssay.loc[(mean1 > 0.3) & (mean3 > 1), "AU_CUT"] = p20 * 0.7 
    filtredFileAssay.loc[(mean1 > 0.3) & (p20 < 0.3), "AU_CUT"] = 0.3
    filtredFileAssay.loc[filtredFileAssay["AU_CUT"].isna(), "AU_CUT"] = p20

    exportFileXlsx(filtredFileAssay, "17)Выборка_проб_метод_1_итоговая", index=False)

    # === По 4 скв и 10 % ==

    # Вычисление метрограмма для выборке
    filtredFileAssay["Метрограмм"] = filtredFileAssay["Длина"] * filtredFileAssay["Au, г/т"]


    # Номера проб
    filtredFileAssay["номер пробы"] = filtredFileAssay.groupby("Субблок").cumcount() + 1
    filtredFileAssay["Разбивка по 30"] = ((filtredFileAssay["номер пробы"] - 1) // 30) + 1

    # Расчёт суммы метрограмма и длин проб на каждую 30-ку
    sumForAllThirty = filtredFileAssay.copy()
    sumForAllThirty["счёт"] = (
        sumForAllThirty.groupby(["Субблок", "Разбивка по 30"])["Разбивка по 30"]
        .transform("count")
    )
    sumForAllThirty["Длина"] = pd.to_numeric(sumForAllThirty["Длина"], errors="coerce")
    sumForAllThirty["Метрограмм"] = pd.to_numeric(sumForAllThirty["Метрограмм"], errors="coerce")

    sumForAllThirty["Метрограмм"] = (
        sumForAllThirty.groupby(["Субблок", "Разбивка по 30"])["Метрограмм"]
        .transform("sum")
    )
    sumForAllThirty["Длина"] = (
        sumForAllThirty.groupby(["Субблок", "Разбивка по 30"])["Длина"]
        .transform("sum")
    )
    sumForAllThirty = sumForAllThirty.drop_duplicates(subset=["Субблок", "Разбивка по 30"])

    sumForAllThirty["Ураган"] = (
        sumForAllThirty.groupby(["Субблок", "Разбивка по 30"])["Метрограмм"]
        .transform("sum") * 0.1
    )

    sumForAllThirty["Суммарная длина"] = (
        sumForAllThirty.groupby(["Субблок", "Разбивка по 30"])["Длина"]
        .transform("sum")
    )
    exportFileXlsx(sumForAllThirty, "18)Суммы по 30-кам", index=False)

    # Объединяем DataFrame по колонкам "Субблок" и "Разбивка по 30"

    colsToAdd = ["Суммарная длина", "Ураган"]
    filtredFileAssay = filtredFileAssay.drop(columns=colsToAdd, errors="ignore")  # удаляем старые
    filtredFileAssay = filtredFileAssay.merge(
        sumForAllThirty[["Субблок", "Разбивка по 30"] + colsToAdd],
        on=["Субблок", "Разбивка по 30"],
        how="left"
    )

    # Маркировка урагана
    for index, row in filtredFileAssay.iterrows():
        metrogram = row["Метрограмм"]
        tallGradeAu = row["Ураган"]
        gradeAu = row["Au, г/т"]

        if metrogram > tallGradeAu and gradeAu >= 1:
            filtredFileAssay.loc[index, "Метка урагана"] = 1
        elif metrogram > tallGradeAu and gradeAu < 1:
            filtredFileAssay.loc[index, "Метка урагана"] = 2
        else:
            filtredFileAssay.loc[index, "Метка урагана"] = 3


    MASK_HORIZONT = fileAssay["Горизонт"] == EXPLORATION_BLOCK.split("_")[0]
    # понижаем в файле композитов по горизонту содержание
    fileAssay.loc[MASK_HORIZONT, "AU 2.5"] = fileAssay.loc[MASK_HORIZONT, "Au, г/т"].clip(upper=2.5)

    radiusFind = 11

    logging.info(f"Радиус поиска: {radiusFind} метров.")

    holeCoords = {
    row["Скважина"]: (float(row["X"]), float(row["Y"]))
    for _, row in filtredFileAssay.iterrows()
    }
    nearby_sorted = {}
    # Для каждой строки из filtered_metrogram — ищем ближайшие скважины из всего набора
    for _, row in fileAssay.iterrows():
        currentHole = row["Скважина"]
        horizont = row["Горизонт"]
        x0 = float(row["X"])
        y0 = float(row["Y"])

        # Находим ближайшие, исключая саму скважину
        if int(horizont) == int(EXPLORATION_BLOCK.split("_")[0]):
            nearby = []
            for other_hole, (x, y) in holeCoords.items():
                if other_hole == currentHole:
                    continue
                dx = x0 - x
                dy = y0 - y
                distance = (dx**2 + dy**2)**0.5
                if distance <= radiusFind:
                    nearby.append(other_hole)

            # Сортируем по расстоянию
            nearby_sorted[currentHole] = nearby[:4]

            rowsNearby = []
            for main_hole, neighbors in nearby_sorted.items():
                for idx, neighbor in enumerate(neighbors, start=1):
                    rowsNearby.append({
                        "Скважина": main_hole,
                        "Ближайшая_скважина": neighbor,
                        "Номер_по_порядку": idx
                    })

            dfNearby = pd.DataFrame(rowsNearby)

            near_hole_dict = defaultdict(list)
            for i, row in dfNearby.iterrows():
                hole = row["Скважина"]
                near_hole = row["Ближайшая_скважина"]
                if len(near_hole) >= 2:
                    near_hole_dict[hole].append(near_hole)
                else:
                    continue

            for hole, near_holes in near_hole_dict.items():
                logging.debug(f"Скважина = {hole} => ближайшие 4-е скважины: {near_holes}")

    logging.info(f"Поиск скважин закончен!")

    for numberHole, nearbyNumberHole in near_hole_dict.items():
        # ищем строку по ключевой скважине
        mask = filtredFileAssay["Скважина"] == numberHole

        if not mask.any():
            logging.warning(f"Скважина {numberHole} не найдена в DataFrame. Пропускаем.")
            continue

        # проверяем метку урагана
        mark_value = filtredFileAssay.loc[mask, "Метка урагана"].values[0]
        if mark_value == 1:
            nearbyGradeAuvalues = fileAssay.loc[
                fileAssay["Скважина"].isin(nearbyNumberHole), "Au, г/т"
            ]
            if not nearbyGradeAuvalues.empty:
                avg_value = nearbyGradeAuvalues.mean()
                filtredFileAssay.loc[mask, "AU_4скв"] = avg_value
                logging.debug(
                    f"Скважина {numberHole}: записано среднее содержание золота "
                    f"по {len(nearbyGradeAuvalues)} соседям = {avg_value:.3f}"
                )
            else:
                logging.debug(f"Скважина {numberHole}: у соседей нет данных по золоту")

    # создаём столбец "Метрограмм_CUT", копируем значения из "Метрограмм"
    filtredFileAssay["Метрограмм_CUT"] = filtredFileAssay["Метрограмм"]

    # где "Метка урагана" == 2, переписываем значениями из "Ураган"
    filtredFileAssay.loc[filtredFileAssay["Метка урагана"] == 2, "Метрограмм_CUT"] = \
        filtredFileAssay.loc[filtredFileAssay["Метка урагана"] == 2, "Ураган"]
    
    filtredFileAssay.loc[filtredFileAssay["Метка урагана"] == 1, "Метрограмм_CUT"] = \
        filtredFileAssay.loc[filtredFileAssay["Метка урагана"] == 1, "AU_4скв"] \
              * filtredFileAssay.loc[filtredFileAssay["Метка урагана"] == 1, "Длина"]
    

    # Расчёт суммы метрограмма_cut для субблоков
    sumForAllThirty = (
        filtredFileAssay.groupby("Субблок", as_index=False)
        .agg({
            "Метрограмм_CUT": "sum",
            "Длина": "sum"
        })
    )
    sumForAllThirty = sumForAllThirty.drop_duplicates(subset=["Субблок"])

    sumForAllThirty["AU_Ураган"] = (
        sumForAllThirty["Метрограмм_CUT"] / sumForAllThirty["Длина"]
    )

    # создаём словарь субблок -> AU_Ураган
    au_dict = dict(zip(sumForAllThirty["Субблок"], sumForAllThirty["AU_Ураган"]))

    au_series = filtredFileAssay["Субблок"].map(au_dict)

    # обновляем только там, где AU_CUT пустой (NaN)
    filtredFileAssay["AU_CUT"] = filtredFileAssay["AU_CUT"].fillna(au_series)
    
    exportFileXlsx(sumForAllThirty, "19)Сумма_метрограмма_CUT", index=False)
    exportFileXlsx(filtredFileAssay, "Выборка_проб", index=False)
    fileAssay.set_index("№ пробы полевой", inplace=True)
    filtredFileAssay.set_index("№ пробы полевой", inplace=True)

    # обновляем только пересечения по индексу и только указанные столбцы
    fileAssay.update(filtredFileAssay[["Метрограмм", "Ураган", "Суммарная длина", "AU_CUT"]])

    fileAssay.reset_index(inplace=True)
    exportFileXlsx(fileAssay, "20)Композиты_СЭР_5м", index=False)

    # создаём словарь Субблок -> AU_Ураган
    au_dict = sumForAllThirty.set_index("Субблок")["AU_Ураган"].to_dict()

    # обновляем файл контуров по Субблоку
    fileStrings["AU_CUT"] = fileStrings["Субблок"].map(au_dict)


    exportFileXlsx(fileStrings, f"21){EXPLORATION_BLOCK}_strings", index=False)


    











cutGrade(r"data\Композиты_CЭР_5м.XLSX", r"data\950_6.5-10.XLSX", EXPLORATION_BLOCK="950_6.5-10")