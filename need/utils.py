import datetime

def get_current_time():
    data = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
    data = "[" + data + " R] "
    return data

def get_current_name():
    data = datetime.datetime.now().strftime('%Y-%m-%d-%H-%M-%S-%f')
    return data

def get_current_date():
    data = datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S')
    return data

def get_current_hour():
    data = datetime.datetime.now().strftime('%H.%M.%S')
    return data