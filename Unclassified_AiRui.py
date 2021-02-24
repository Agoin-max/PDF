from AirRui import AirModel


class UnAirModel(AirModel):

    def time_strappend(self, file):
        path_list = file.rsplit("-", 2)
        return path_list[0] + "-" + self.time_strftime() + ".xlsx"
