from configparser import ConfigParser


def get_config(category, item):
    config = ConfigParser()
    config.read("C:/Users/manju/PycharmProjects/Data_Validation/configuration/config.ini")
    return config.get(category, item)


