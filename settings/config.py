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
                        "Font": {"name": "Calibri", "size": 12, "bold": True,
                                 "color": "45455d"
                                 },
                        "Border": {"border_style": 'medium', "color": '444444'
                                   }
                        },
                    "Sub_Header": {
                        "Alignment": {"horizontal": 'center', "vertical": 'center'
                                      },
                        "Font": {"name": "Calibri Light", "size": 11, "bold": False,
                                 "color": "45455d"
                                 }
                        },
                    "Var_Header": {
                        "Alignment": {"horizontal": 'center', "vertical": 'center'
                                      },
                        "Font": {"name": "Calibri Light", "size": 11, "bold": True,
                                 "color": "45455d"
                                 },
                        "Border": {"border_style": 'medium', "color": '444444'
                                   },
                        "PatternFill": {"start_color": "C5D9F1",
                                        "end_color": "C5D9F1",
                                        "fill_type": "solid"}
                        },
                    "Body": {
                            "Font": {"color": "45455d", "name": "Calibri",
                                     "size": 10},
                            "Alignment": {"horizontal": 'center',
                                          "vertical": 'center'}
                            },
                    "Footer": {
                        "Font": {"color": "45455d",
                                 "name": "Calibri",
                                 "size": 8},
                        "Alignment": {"horizontal": 'left',
                                      "vertical": 'center'},
                        "Border": {"border_style": 'medium',
                                   "color": '444444'}
                        }
                    }
                with open(self.__path, "w") as file:
                    json.dump(self.file, file, indent=4)


config_json = JsonFile().file
