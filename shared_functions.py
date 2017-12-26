from re import split

def file_name_split(file_name):
    name_elements = split(r'[\/\-\s]+', file_name.replace('.xlsx', ''))
    return name_elements

def reformat_file_name(file_name, tryout_string='tryout'):
    is_tryouts = True if tryout_string in file_name.lower() else False
    name_elements = file_name_split(file_name)
    if is_tryouts:
        date = name_elements[:-1]
    else:
        date = name_elements
    date_short = (date[0], date[1], date[2][:2])
    date_long = tuple([int(element) for element in (date[0], date[1], '20' + date[2][:2])])
    return date_long, date_short, is_tryouts