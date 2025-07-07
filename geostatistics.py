import json

class GeoStatistics:
    """
    Класс для операций с гео.данными
    """
    def __init__(self, data: dict):
        self.data = data

    def metrogram(self, field_sample_length: str, field_sample_grade:str) -> dict:
        """
        Данные для данного метода в виде данные в
        виде словаря dict{key:[list]}

        Вычисление метрограммы: длина пробы × содержание Au
        :return: словарь {index: метрограмма}
        """
        lengths = self.data[field_sample_length]
        grades = self.data[field_sample_grade]

        if not lengths or not grades:
            raise ValueError("Указанные поля не найдены в данных")

        if len(lengths) != len(grades):
            raise ValueError("Поля должны иметь одинаковую длину")

        return {index: float(lengths[index]) * float(grades[index]) for index in sorted(range(len(lengths)))}
    

    def average(self, field_sample_length: str, field_sample_metrogram: str) -> float:
        """
        Вычисление среднего содержание: сумма длин проб / сумма метрограмма проб
        """
        lengths = self.data[field_sample_length]
        metrogram = self.data[field_sample_metrogram]

        if not lengths or not metrogram:
            raise ValueError("Указанные поля не найдены в данных")

        if len(lengths) != len(metrogram):
            raise ValueError("Поля должны иметь одинаковую длину")

        sum_lengths = 0
        sum_metrogram = 0

        for i in range(len(lengths)):
            sum_lengths += float(lengths[i])
            sum_metrogram += float(metrogram[i])

        return sum_lengths / sum_metrogram

class Strings:
    def __init__(self):
        pass

    @staticmethod
    def is_point_in_polygon(x: float, y: float, polygon: list):
        """
        Проверка, находится ли точка (x, y) внутри полигона.
        polygon — список кортежей (x, y).
        """
        inside = False
        n = len(polygon)
        
        px1, py1 = polygon[0]
        for i in range(1, n + 1):
            px2, py2 = polygon[i % n]
            if y > min(py1, py2):
                if y <= max(py1, py2):
                    if x <= max(px1, px2):
                        if py1 != py2:
                            xinters = (y - py1) * (px2 - px1) / (py2 - py1 + 1e-10) + px1
                        if px1 == px2 or x <= xinters:
                            inside = not inside
            px1, py1 = px2, py2

        return inside
    

class Find:

    def __init__(self):
        pass

    def find_nearby_within_radius(self, center_x, center_y, all_coords, radius):
        """Возвращает список (индекс, расстояние) до всех точек в радиусе"""
        results = []
        for i, (x, y) in enumerate(all_coords):
            dx = float(center_x) - float(x)
            dy = float(center_y) - float(y)
            distance = (dx**2 + dy**2)**0.5
            if 0 < distance <= radius:  # исключаем саму точку (0)
                results.append((i, distance))
        return sorted(results, key=lambda x: x[1])





