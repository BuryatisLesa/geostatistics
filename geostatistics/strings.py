
class Strings:
    def __init__(self):
        pass
    @staticmethod
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
    @staticmethod
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
    
    @staticmethod
    def is_point_in_polygon_with_tol(x, y, polygon, tol=0.2):
        """Проверка: внутри полигона или рядом с границей"""
        if Strings.is_point_in_polygon(x, y, polygon):
            return True
        n = len(polygon)
        for i in range(n):
            x1, y1 = polygon[i]
            x2, y2 = polygon[(i + 1) % n]
            dist = Strings.point_to_segment_distance(x, y, x1, y1, x2, y2)
            if dist <= tol:
                return True
        return False
