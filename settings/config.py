import os
import json


class JsonFile:
    def __init__(self, path="settings/config_style.json"):
        self.__path = path
        self.file = {}
        if os.path.exists(self.__path):
            with open(self.__path, "r") as file:
                self.file = json.load(file)
        else:
            if self.__path == "settings/config_style.json":
                self.file = {
                    "Header": {
                        "Alignment": {"horizontal": 'left', "vertical": 'center'
                                      },
                        "Font": {"name": "Gotham Book", "size": 12, "bold": True,
                                 "color": "45455d"
                                 },
                        "Border": {"border_style": 'medium', "color": '86754d'
                                   }
                        },
                    "Sub_Header": {
                        "Alignment": {"horizontal": 'center', "vertical": 'center'
                                      },
                        "Font": {"name": "Gotham Light", "size": 11, "bold": False,
                                 "color": "45455d"
                                 }
                        },
                    "Var_Header": {
                        "Alignment": {"horizontal": 'center', "vertical": 'center'
                                      },
                        "Font": {"name": "Gotham Light", "size": 11, "bold": True,
                                 "color": "45455d"
                                 },
                        "Border": {"border_style": 'medium', "color": '86754d'
                                   },
                        "PatternFill": {"start_color": "c3b6a1",
                                        "end_color": "c3b6a1",
                                        "fill_type": "solid"}
                        },
                    "Body": {
                            "Font": {"color": "45455d", "name": "Gotham Medium",
                                     "size": 10},
                            "Alignment": {"horizontal": 'center',
                                          "vertical": 'center'}
                            },
                    "Footer": {
                        "Font": {"color": "45455d",
                                 "name": "Gotham Extra Light",
                                 "size": 8},
                        "Alignment": {"horizontal": 'left',
                                      "vertical": 'center'},
                        "Border": {"border_style": 'medium',
                                   "color": '86754d'}
                        }
                    }
                with open(self.__path, "w") as file:
                    json.dump(self.file, file, indent=4)


config_json = JsonFile().file
