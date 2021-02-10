# -*- coding: utf-8 -*-

import time
import os

def time_track(func):
    def surrogate(*args, **kwargs):
        started_at = time.time()

        result = func(*args, **kwargs)

        ended_at = time.time()
        elapsed = round(ended_at - started_at, 4)
        print(f'Функция {func.__name__} работала {elapsed} секунд(ы) или {round(elapsed/60, 1)} минут(ы)')
        return result
    return surrogate

def criate_log_file(file_name=None):
    if file_name is None:
        f_name = 'function_errors.log'
    else:
        f_name = os.path.normpath(file_name)
    def log_errors(func):
        def surrogate(*args, **kwargs):
            try:
                result = func(*args, **kwargs)
            #  Здесь можно перехватывать только Exception, а тип исключения
            #  брать из exc. Это позволит убрать дублирование кода.
            except Exception as exc:
                file = open(f_name, 'a', encoding='UTF-8')
                error_str = f'{str(func)} {type(exc)} {exc} \n'
                file.write(error_str)
                file.close()
                raise exc
            return result
        return surrogate
    return log_errors