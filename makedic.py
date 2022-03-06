def make_class_dic(df, year, start_month):
    condition1 = df['year'] == year
    condition2 = df['start_month'] == start_month
    df_this_month = df[condition1 & condition2]
    class_dic = {}
    for class_name, class_addr in zip(df_this_month['name'], df_this_month['address']):
        class_dic[class_name] = f'https://band.us/band/{int(class_addr)}'
    return class_dic
    