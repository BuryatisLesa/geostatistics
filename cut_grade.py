import pandas as pd
from geostatistics.geostatistics import GeoStatisctics as gs
import logging
import os
from geostatistics.strings import Strings
from collections import defaultdict

def cutGrade(pathFileAssay, pathFileStrings, EXPLORATION_BLOCK):

    logging.basicConfig(
        level=logging.DEBUG,
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
    exportFileXlsx(filtredFileAssay, "Выборка_проб", index=False)

    # Номера проб

    print(filtredFileAssay.groupby("Субблок")["ЧАСТОТА"])












cutGrade(r"data\Композиты_CЭР_5м.XLSX", r"data\950_6.5-10.XLSX", EXPLORATION_BLOCK="950_6.5-10")