import pandas as pd

class GeoStatisctics:
    def __init__(self):
        pass


    def filteredData(file, field, filter, operator="contains"):
        if operator == "contains":
            mask = file[field] == filter
            filtredFile = file[mask]
            return filtredFile





