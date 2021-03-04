#! coding=utf8

from os.path import exists
from configparser import ConfigParser


class Config:
    """
        @requires:
                log-- an object for writing log
                config_path-- path of config file
        @return:
                a dictionary of data
        Read the config file and return
    """
    def __init__(self, config_path):
        self.config_path = config_path
        self.config = ConfigParser()

    def read(self):
        dic = {}
        if not exists(self.config_path):
            raise IOError('No found config file')

        self.config.read(self.config_path)
        # Get all sections and keys
        for section in self.config.sections():
            for key in self.config.options(section):
                dic[key] = self.config.get(section, key)
        return dic

    def write(self, data):
        open(self.config_path, 'w').close()
        self.config.read(self.config_path)
        self.config.add_section('signature')
        for key, value in data.items():
            if isinstance(value, list) and len(value) > 1:
                value = ','.join(iter(value))
            self.config.set('signature', key, value)
        self.config.write(open(self.config_path, 'w'))

    def get_data(self, section, key):
        """
        直接获取指定section下，指定key的数据
        :param section:
        :param key:
        :return:
        """
        # if not exists(self.config_path):
        #     raise FileNotFoundError('No found config file')
        self.config.read(self.config_path)

        if section not in self.config.sections():
            raise ValueError('No found section:'+section)
        if key not in self.config.options(section):
            raise ValueError('No found key:'+key)

        return self.config.get(section, key)
