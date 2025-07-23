from typing import Tuple, Dict, List, Union
from collections import defaultdict
from data import all_points

def calculate_distance_point(
        point1: Tuple[float, float],
        point2: Tuple[float, float]
        ) -> float:
    """Функция по вычисление расстояние между точками"""
    
    x1, y1 = point1
    x2, y2 = point2
    dx = x2 - x1
    dy = y2 - y1
    distance = (dx**2 + dy**2)**0.5
    return distance

def coordinates_mid_points(
        point1: Tuple[float, float],
        point2: Tuple[float, float]
        ) -> Tuple[float, float]:
    """ Вычисление координат половина расстояния между точками """

    x1, y1 = point1
    x2, y2 = point2
    x = (x1 + x2) / 2
    y = (y1 + y2) / 2
    coords_mid_point = (x, y)
    return coords_mid_point


def find_point(
        all_points: Dict[int, Tuple[float, float]],
        get_coords_nearby_points = False
        ) -> Union[
            Dict[int, List[Tuple[int, float]]],
            Dict[int, List[Tuple[int, Tuple[float, float]]]]]:
    """Найти ближайшие четыре точки"""

    four_point = defaultdict()
    for number_point, current_coords_point in all_points.items():
        distance_point = defaultdict()

        for point, coords in all_points.items():
            distance_point[point] = calculate_distance_point(
                current_coords_point, coords)
        
        filter_min_distance = sorted(
            distance_point.items(),
            key=lambda item: item[1])[1:5]
        
        four_point[number_point] = filter_min_distance
    
    if not get_coords_nearby_points:
         return four_point
    
    four_point_with_coords = defaultdict(list)
    for current_point , arr_nearby_points in four_point.items():
        for nearby_point, distance_point in arr_nearby_points:
            for point, coords in all_points.items():
                if nearby_point == point:
                    four_point_with_coords[current_point].append((nearby_point, coords))

    return four_point_with_coords