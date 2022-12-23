from lib_main import *

def final_compilate(purchase_book, 
                    archivebook,
                    categorybook,
                    orderbook,
                    archmybook,
                    collisionbook,
                    e: excel.Excel, archdict: dict,
                    archivedick: dict,
                    purchase_one_dict: dict,
                    purchase_two_dict: dict,
                    my_mail: mail.Attachmail):

    """MainInformation."""

    book_check = e.load_book(dir=config.BOOK_CHECK)
    file = e.findexc(dir=config.ROOT_DIR)
    print(file, 'YA FILE')
    try:

        mybook = op.load_workbook(filename=file)

    except:

        os.remove(file)
        #body_messages = (f"""Письмо с темой: {my_mail.R_SUBJECT}. 
#Не подходит для обработки, 
#скорее всего в письме нет нужного формата файла!:(
#""")

        #my_mail.send_email(user=config.USERNAME_GMAIL,
                             #pwd=config.PASSWORD_GMAIL,
                             #recipient=f'{my_mail.R_EMAIL}',
                             #subject=f'{my_mail.R_SUBJECT}',
                             #body=body_messages,
                             #file=None,
                             #filename=None)

        return False

    mysheet = mybook.worksheets
    finaldict: dict = {}
    count: int = 0
    list_item, sheet = e.isParametr(sheetname=mysheet)

    if sheet is None:

        if len(list_item) == 0:

            body_messages = (f"""Письмо с темой: {my_mail.R_SUBJECT}.
Файл {e.MAIN_FILENAME} не подходит для обработки. 
Нет нужных столбцов!:(
""")

        os.remove(file)
        my_mail.send_email(user=config.USERNAME_GMAIL,
                             pwd=config.PASSWORD_GMAIL,
                             recipient=f'{my_mail.R_EMAIL}',
                             subject=f'{my_mail.R_SUBJECT}',
                             body=body_messages,
                             file=None,
                             filename=None)

        return False

    main_list: list = []

    """StartMainDriver."""
    for keys_main in list_item:

        flag: bool = False
        count += 1
        local_item: dict = {}
        item: dict = {}
        local_item_atr: dict = {}
        ws = (book_check['Расчет'], book_check['Для архива'])

        item['Amount'] = keys_main[1]

        if keys_main[2] != 'None':

            category_desc = e.categorydescription(param=keys_main, 
                                                 book=collisionbook, 
                                                 category_book=categorybook)

            if category_desc is not False:
                item.update(category_desc)
                local_item_atr = item.copy()

        keys_upper = e.filterkey(key=keys_main[0])
        listkey, num = e.collision(key=keys_main[0], 
                                   vendor=keys_main[3])
        
        for rowa in range(0, num):
            if flag is True:
                break

            if len(listkey) > 1:
                keys = listkey[rowa]
            else:
                keys = listkey[0]
            if keys in archdict:

                item.update(archdict[keys])
                if 'Category' not in item:
                    item.update(e.archivefinding(local_key=keys, 
                                                archive_dict=archivedick, 
                                                category_book=categorybook, 
                                                main_key=keys_main))

                item.update(e.nowritercatebay(item=item, 
                                              key=keys_main[0]))

                finaldict[keys_upper] = item
                local_item[keys_upper] = item
                main_list.append(local_item)
                e.writing(item_dict=item, 
                         ws=ws, 
                         book='WriteZIP', 
                         appoint='No', 
                         event='None', 
                         l_count=count, 
                         localKey='None', 
                         mainkey=keys_main,
                         list_item=list_item)

                flag = True

        if flag is True:
            continue

        if flag is False:
            abstract_len_in = 0
            abstract_len_out = 0
            arch_key_in = 0
            arch_key_out = 0
            my_list = list(archdict.keys())
            for key_arch in my_list:
                if key_arch is None:
                    continue
                if keys_upper in key_arch:
                    if len(key_arch) > abstract_len_in:
                        abstract_len_in = len(key_arch)
                        arch_key_in = key_arch
                        flag = True

                elif key_arch in keys_upper:
                    if len(key_arch) > abstract_len_out:
                        abstract_len_out = len(key_arch)
                        arch_key_out = key_arch
                        flag = True

            if flag is True:
                if arch_key_in != 0:
                    item.update(archdict[arch_key_in])
                    if 'Category' not in item:

                        item.update(e.archivefinding(local_key=arch_key_in, 
                                                     archive_dict=archivedick, 
                                                     category_book=categorybook, 
                                                     main_key=keys_main))

                    item.update(e.nowritercatebay(item=item, 
                                                  key=keys_main[0]))

                    finaldict[keys_upper] = item
                    local_item[keys_upper] = item
                    main_list.append(local_item)
                    e.writing(item_dict=item, 
                             ws=ws, 
                             book='WriteZIP', 
                             appoint='No', 
                             event='Collision', 
                             l_count=count, 
                             localKey=keys_main[0], 
                             mainkey=keys_main,
                             list_item=list_item)

                else:
                    item.update(archdict[arch_key_out])
                    if 'Category' not in item:
                        item.update(e.archivefinding(local_key=arch_key_out, 
                                                    archive_dict=archivedick, 
                                                    category_book=categorybook,
                                                    main_key=keys_main))

                    item.update(e.nowritercatebay(item=item, key=keys_main[0]))

                    finaldict[keys_upper] = item
                    local_item[keys_upper] = item
                    main_list.append(local_item)                    
                    e.writing(item_dict=item, 
                             ws=ws, 
                             book='WriteZIP', 
                             appoint='No', 
                             event='Collision', 
                             l_count=count, 
                             localKey=keys_main[0], 
                             mainkey=keys_main,
                             list_item=list_item)

        if flag is True:
            continue

        if flag is False:
            for rowa in range(0, num):
                if flag is True:
                    break

                if len(listkey) > 1:
                    keys = listkey[rowa]
                else:
                    keys = listkey[0]

                if keys in purchase_one_dict:
                    item.update(purchase_one_dict[keys])
                    if 'Category' not in item:
                        item.update(e.archivefinding(local_key=keys, 
                                                     archive_dict=archivedick, 
                                                     category_book=categorybook, 
                                                     main_key=keys_main))

                    item.update(e.nowritercatebay(item=item, 
                                                  key=keys_main[0]))
                    
                    finaldict[keys_upper] = item
                    local_item[keys_upper] = item
                    main_list.append(local_item)
                    e.writing(item_dict=item, 
                              ws=ws, 
                              appoint='48port', 
                              book='WriteZIP', 
                              event='None', 
                              l_count=count, 
                              localKey='None', 
                              mainkey=keys_main,
                              list_item=list_item)

                    flag = True

        if flag is True:
            continue
        
        if flag is False:
            for key_purchase_one in purchase_one_dict.keys():
                if flag is True:
                    break
                if key_purchase_one is None:
                    continue

                if keys_upper in str(key_purchase_one):
                    item.update(purchase_one_dict[key_purchase_one])
                    if 'Category' not in item:
                        item.update(e.archivefinding(local_key=key_purchase_one, 
                                                     archive_dict=archivedick, 
                                                     category_book=categorybook, 
                                                     main_key=keys_main))

                    item.update(e.nowritercatebay(item=item, 
                                                  key=keys_main[0]))
                    
                    finaldict[keys_upper] = item
                    local_item[keys_upper] = item
                    main_list.append(local_item)
                    e.writing(item_dict=item, 
                              ws=ws, 
                              appoint='48port', 
                              book='WriteZIP', 
                              event='Collision', 
                              l_count=count, 
                              localKey=keys_main[0], 
                              mainkey=keys_main,
                              list_item=list_item)

                    flag = True

                elif str(key_purchase_one) in keys_upper:

                    item.update(purchase_one_dict[key_purchase_one])
                    if 'Category' not in item:
                        item.update(e.archivefinding(local_key=key_purchase_one, 
                                                     archive_dict=archivedick, 
                                                     category_book=categorybook, 
                                                     main_key=keys_main))

                    item.update(e.nowritercatebay(item=item, 
                                                  key=keys_main[0]))

                    finaldict[keys_upper] = item
                    local_item[keys_upper] = item
                    main_list.append(local_item)
                    e.writing(item_dict=item, 
                              ws=ws, 
                              appoint='48port', 
                              book='WriteZIP', 
                              event='Collision', 
                              l_count=count, 
                              localKey=keys_main[0], 
                              mainkey=keys_main,
                              list_item=list_item)

                    flag = True

        if flag is True:
            continue

        if flag is False:
            for rowa in range(0, num):
                if flag is True:
                    break

                if len(listkey) > 1:
                    keys = str(listkey[rowa])
                else:
                    keys = str(listkey[0])
                if keys in purchase_two_dict:
                    item.update(purchase_two_dict[keys])
                    if 'Category' not in item:
                        item.update(e.archivefinding(local_key=keys, 
                                    archive_dict=archivedick, 
                                    category_book=categorybook, 
                                    main_key=keys_main))

                    item.update(e.nowritercatebay(item=item, 
                                                  key=keys_main[0]))

                    finaldict[keys_upper] = item
                    local_item[keys_upper] = item
                    main_list.append(local_item)
                    e.writing(item_dict=item, 
                              ws=ws, 
                              appoint='48port', 
                              book='WriteZIP', 
                              event='None', 
                              l_count=count, 
                              localKey='None', 
                              mainkey=keys_main,
                              list_item=list_item)

                    flag = True

        if flag is True:
            continue

        if flag is False:
            for key_purchase_two in purchase_two_dict.keys():
                if flag is True:
                    break
                if key_purchase_two is None:
                    continue

                if keys_upper in str(key_purchase_two):
                    item.update(purchase_two_dict[key_purchase_two])
                    if 'Category' not in item:
                        item.update(e.archivefinding(local_key=key_purchase_two, 
                                                     archive_dict=archivedick, 
                                                     category_book=categorybook, 
                                                     main_key=keys_main))

                    item.update(e.nowritercatebay(item=item, key=keys_main[0]))

                    finaldict[keys_upper] = item
                    local_item[keys_upper] = item
                    main_list.append(local_item)
                    e.writing(item_dict=item, 
                              ws=ws, 
                              appoint='48port', 
                              book='WriteZIP', 
                              event='Collision', 
                              l_count=count, 
                              localKey=keys_main[0], 
                              mainkey=keys_main,
                              list_item=list_item)

                    flag = True

                elif str(key_purchase_two) in keys_upper:
                    item.update(purchase_two_dict[key_purchase_two])
                    if 'Category' not in item:
                        item.update(e.archivefinding(local_key=key_purchase_two, 
                                                    archive_dict=archivedick, 
                                                    category_book=categorybook, 
                                                    main_key=keys_main))

                    item.update(e.nowritercatebay(item=item, key=keys_main[0]))

                    finaldict[keys_upper] = item
                    local_item[keys_upper] = item
                    main_list.append(local_item)
                    e.writing(item_dict=item, 
                              ws=ws, 
                              appoint='48port', 
                              book='WriteZIP', 
                              event='Collision', 
                              l_count=count, 
                              localKey=keys_main[0], 
                              mainkey=keys_main,
                              list_item=list_item)

                    flag = True

        if flag is True:
            continue

        if flag is False:
            for rowa in range(0, num):
                if flag is True:
                    break

                if len(listkey) > 1:
                    keys = str(listkey[rowa])
                else:
                    keys = str(listkey[0])

                if keys in archivedick:
                    item.update(archivedick[keys])
                    final_category = e.categoryfinal(param=item, 
                                                     categoryybook=categorybook)

                    item.update(final_category)
                    if 'Category' in local_item_atr:
                        item.update(local_item_atr)

                    item.update(e.nowritercatebay(item=item, 
                                                  key=keys_main[0]))

                    finaldict[keys_upper] = item
                    local_item[keys_upper] = item
                    main_list.append(local_item)
                    e.writing(item_dict=item, 
                              ws=ws, 
                              appoint='Archive', 
                              book='WriteZIP', 
                              event='None', 
                              l_count=count, 
                              localKey='None', 
                              mainkey=keys_main,
                              list_item=list_item)

                    flag = True

        if flag is True:
            continue

        if flag is False:
            for key_archive in archivedick.keys():
                if flag is True:
                    break
                if key_archive is None:
                    continue
                if keys_upper in str(key_archive):
                    item.update(archivedick[key_archive])

                    item.update(e.categoryfinal(param=item, 
                                categoryybook=categorybook))

                    if 'Category' in local_item_atr:
                        item.update(local_item_atr)

                    item.update(e.nowritercatebay(item=item, 
                                                  key=keys_main[0]))

                    finaldict[keys_upper] = item
                    local_item[keys_upper] = item
                    main_list.append(local_item)
                    e.writing(item_dict=item, 
                              ws=ws, 
                              appoint='Archive', 
                              book='WriteZIP', 
                              event='Collision', 
                              l_count=count, 
                              localKey=keys_main[0], 
                              mainkey=keys_main,
                              list_item=list_item)

                    flag = True

                elif str(key_archive) in keys_upper:

                    item.update(archivedick[key_archive])
                    item.update(e.categoryfinal(param=item, 
                                                categoryybook=categorybook))
                    if 'Category' in local_item_atr:
                        item.update(local_item_atr)

                    item.update(e.nowritercatebay(item=item, 
                                                  key=keys_main[0]))

                    finaldict[keys_upper] = item
                    local_item[keys_upper] = item
                    main_list.append(local_item)
                    e.writing(item_dict=item, 
                              ws=ws, 
                              appoint='Archive', 
                              book='WriteZIP', 
                              event='Collision', 
                              l_count=count, 
                              localKey=keys_main[0], 
                              mainkey=keys_main,
                              list_item=list_item)

                    flag = True

        if flag is True:
            continue

        if flag is False:
            try:
                if 'Category' not in item:
                    item.update(e.categoryfinal(param=e.categoryarch(key=keys_upper, 
                                                                     categorybook=categorybook), 
                                                                     categoryybook=categorybook))

                item.update(e.nowritercatebay(item=item, 
                                              key=keys_main[0]))
            except:
                item.update(e.nowritercatebay(item=item, 
                                              key=keys_main[0]))

            finaldict[keys_upper] = item
            local_item[keys_upper] = item
            main_list.append(local_item)
            e.writing(item_dict=item, 
                      ws=ws, 
                      appoint='Other', 
                      book='WriteZIP', 
                      event='None', 
                      l_count=count, 
                      localKey='None', 
                      mainkey=keys_main,
                      list_item=list_item)

    e.MAIN_FILE_OUT = config.OUT_DIR + '/final.xlsx'
    
    book_check.save(e.MAIN_FILE_OUT)
    os.remove(e.MAIN_FILE_IN)

    json_object_archive = json.dumps(finaldict,
                                     indent=4,
                                     ensure_ascii=False)

    with open("Final.json", "w", encoding='utf-8') as outfile:
        outfile.write(json_object_archive)

    return True


if __name__ == '__main__':
    """Main driver"""

    """MainInformation."""
    e = excel.Excel(api_key=config.API_KEY_EBAY)
    
    """MainBook."""

    """MainBook."""
    purchase_book = e.load_book(dir=config.PURCHASE_BOOK)
    archivebook = e.load_book(dir=config.ARCHIVE_BOOK)
    categorybook = e.load_book(dir=config.CATEGORY_BOOK)
    orderbook = e.load_book(dir=config.ORDER_BOOK)
    archmybook = e.load_book(dir=config.ARCH_BOOK)
    collisionbook = e.load_book(dir=config.COLLISION_BOOK)

    """MainDict."""
    archdict = e.archcreation(archmybook=archmybook)

    archivedick = e.archivecreation(archivebook=archivebook, 
                                    orderbook=orderbook, 
                                    category=categorybook
                                    )

    purchase_one_dict = e.purchasesearch_one(purchasebook=purchase_book)
    purchase_two_dict = e.purchasesearch_two(purchasebook=purchase_book)


    #Start_check_mail
    while (True):
        
        sleep(1)
        my_mail = mail.Attachmail(
                            mail_server=config.MAIL_SERVER,
                            username=config.USERNAME_GMAIL,
                            password=config.PASSWORD_GMAIL
                            )
                            
        process_mail = my_mail.get_attachments(config.ROOT_DIR,
                                               e)
        print(process_mail, 'Процесс)')

        if process_mail is True:

            while (e.findexc(dir=config.ROOT_DIR) != ''):
                
                #Main_func
                result = final_compilate(
                                 purchase_book=purchase_book,
                                 archivebook=archivebook,
                                 categorybook=categorybook,
                                 orderbook=orderbook,
                                 archmybook=archmybook,
                                 collisionbook=collisionbook,
                                 e=e, archdict=archdict,
                                 archivedick=archivedick,
                                 purchase_one_dict=purchase_one_dict,
                                 purchase_two_dict=purchase_two_dict,
                                 my_mail=my_mail)

                if result is False:
                    continue

                file_out = e.MAIN_FILE_OUT
                filename = e.MAIN_FILENAME
                max_count = e.MAX_COUNT

                wbxl = xw.Book(e.MAIN_FILE_OUT)

                multiplying = wbxl.sheets['Оценка рыночной стоимости'].range('B2').value
                lower_bound = wbxl.sheets['Оценка рыночной стоимости'].range('B3').value
                upper_bound = wbxl.sheets['Оценка рыночной стоимости'].range('B4').value
                wbxl.close()

                try:
                    f_lower_bound = '{:,}'.format(math.ceil(lower_bound)).replace(',', ' ')
                    f_upper_bound = '{:,}'.format(math.ceil(upper_bound)).replace(',', ' ')
                except:
                    f_lower_bound = 0
                    f_upper_bound = 0 

  
                table_text = ''
                body_messages = (f"""Обработанный файл во вложении!:)\n""")

                if max_count < 20:

                    disclaimer = ("""\nПозиций очень мало, поэтому повышающему 
коэффициенту верить не рекомендуется. 
Требуется перепроверка доли ненайденных позиций вручную.\n""")

                    body_messages += disclaimer 

                if multiplying > 1.2:
 
                    disclaimer = ("""\nБольшой повышающий коэффициент. 
Требуется перепроверка доли ненайденных позиций вручную.\n""")

                    body_messages += disclaimer

                if lower_bound < upper_bound:

                    table_text = (f"""\n«Мы должны вписываться в диапазон 
от {str(f_lower_bound) + 'р.'}  и до {str(f_upper_bound) + 'р.'} с учетом того, 
что {math.ceil((multiplying - 1) * 100)}% единиц оборудования система не нашла в продаже на рынке.\n
Расчет выполнен по курсу 70р/$»\n""")

                if lower_bound > upper_bound:

                    table_text = (f"""\n«Мы должны вписываться в диапазон 
от {str(f_upper_bound) + 'р.'} и до {str(f_lower_bound) + 'р.'} с учетом того, 
что {math.ceil((multiplying - 1) * 100)}% единиц оборудования система не нашла в продаже на рынке.\n
Расчет выполнен по курсу 70р/$»\n""")

                body_messages += table_text

                #for mail in 
                my_mail.send_email(user=config.USERNAME_GMAIL,
                               pwd=config.PASSWORD_GMAIL,
                               recipient= my_mail.R_EMAIL,
                               subject= my_mail.R_SUBJECT,
                               body=body_messages,
                               file=file_out,
                               filename=filename)

                os.remove(file_out)

            my_mail.close_connection()

