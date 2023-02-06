from lib_excel import *


class Excel:

    MAIN_FILE_OUT: str = None
    MAIN_FILENAME: str = None
    MAIN_SHEET = None
    MYBOOK = None

    MAX_COUNT: int = 0

    def __init__(self, api_key):

        self.api_key_ebay = api_key

    def findexc(self, dir) -> list:
        """search excel file in folder."""
        file_dir: str = ''
        flag: bool = True
        for dirName, subdirList, fileList in os.walk(dir):

            if flag is False:
                break

            for fname in fileList:

                if '.xlsx' in fname or '.xlsm' in fname:

                    file_dir: str = dir + f'/{str(fname)}'
                    self.MAIN_FILENAME = str(fname)
                    flag = False
                    break

        return file_dir

    def load_book(self, dir):

        file = self.findexc(dir=dir)
        book = op.load_workbook(filename=file)
        return book

    def isParametr(self, book):
        """search for the desired column."""

        sheetname = book.worksheets
        final_list: list = []
        column_list: list = []
        flag: bool = False

        #param_point
        name_c: int = 0
        pn_c: int = 0
        amount_c: int = 0
        description_c: int = 0
        vendor_c: int = 0

        #collision
        col = ('NONE', 'None', '-', '')

        for sheet in sheetname:

            count: int = 0
            number_true: int = 0

            for COL in sheet.iter_cols(1, sheet.max_column + 1):
                count += 1

                if number_true == 5:

                    flag = True
                    it_tuple: tuple = (sheet, (name_c, vendor_c, pn_c, description_c, amount_c))
                    column_list.append(it_tuple)
                    break

                if type(COL[0].value) is not str:
                    continue

                column_value = COL[0].value.upper()

                if 'P/N' in column_value or 'ПАРТ-НОМЕР' in column_value:

                    pn_c = count
                    number_true += 1

                elif 'КОЛ-ВО' in column_value or 'КОЛИЧЕСТВО' in column_value:

                    amount_c = count
                    number_true += 1

                elif 'ОПИСАНИЕ' in column_value:

                    description_c = count
                    number_true += 1

                elif 'ВЕНДОР' in column_value:

                    vendor_c = count
                    number_true += 1

                elif 'ЗАКАЗЧИК' in column_value:

                    name_c = count
                    number_true += 1
            
        if flag is False:
            return final_list

        for item in column_list:

            list_item: list = []
            sheet_name = item[0]
            column_number = item[1]

            for rowa in range(2, sheet_name.max_row + 1):
                

                description = str(sheet_name.cell(row=rowa,
                                        column=column_number[3]).value).upper()

                part_number = str(sheet_name.cell(row=rowa,
                                        column=column_number[2]).value).upper()

                amount = sheet_name.cell(row=rowa,
                                    column=column_number[4]).value

                vendor = str(sheet_name.cell(row=rowa,
                                    column=column_number[1]).value).upper()

                name = str(sheet_name.cell(row=rowa,
                                    column=column_number[0]).value).upper()
                
                if part_number is None or part_number in col:
                    continue


                if amount is None or amount in col:
                    amount = 0

                if description is None or description in col:
                    description = ''
                
                item_tuple = (part_number,
                            amount, description, vendor, name)

                list_item.append(item_tuple)

            final_list.append((list_item, sheet_name.max_row, sheet_name))

        return final_list

    def purchasesearch_one(self, purchasebook):
        """purchase creation."""

        p_dict: dict = {}
        sheet = purchasebook['Закупается']

        for rowa in range(2, sheet.max_row):

            pn = str(sheet.cell(row=rowa, column=1).value)
            price_zip = sheet.cell(row=rowa, column=9).value
            service_comment = sheet.cell(row=rowa, column=6).value
            appoint = sheet.cell(row=rowa, column=4).value
            filter_key = self.filterkey(key=pn).upper()

            if filter_key not in p_dict:

                item = {'PriceCostZIP': f'{price_zip}',
                        'ZIP': f'{pn}',
                        'ServiceComment': f'Закупается под: {service_comment}',
                        'Appoint': f'{appoint}',
                        'MainKey': f'{pn}'}

                p_dict[filter_key] = item

        return p_dict

    def purchasesearch_two(self, purchasebook):
        """purchase creation."""

        p_dict: dict = {}
        sheet = purchasebook['Хотим']
        for rowa in range(2, sheet.max_row):

            pn = str(sheet.cell(row=rowa, column=1).value)
            pricezip = sheet.cell(row=rowa, column=3).value
            appoint = sheet.cell(row=rowa, column=8).value
            pricegrabe = sheet.cell(row=rowa, column=4).value

            filter_key = self.filterkey(key=pn).upper()

            if filter_key not in p_dict:

                item = {'PriceCostZIP': f'{pricezip}',
                        'ZIP': f'Есть в планах: {pn}',
                        'Appoint': f'{appoint}',
                        'MainKey': f'{pn}'}

                atr = sheet.cell(row=rowa, column=7).value
                if atr is not None:
                    item['ServiceComment'] = f'''Хотим купить под:
                                             {atr}. По цене:
                                             {pricegrabe}$'''

                else:

                    atr_two = sheet.cell(row=rowa, column=2).value
                    item['ServiceComment'] = f'''Хотим купить под:
                                             {atr_two}. По цене:
                                             {pricegrabe}$'''

                p_dict[filter_key] = item

            return p_dict

    def archcreation(self, archmybook):
        """arch creation."""

        archdict: dict = {}
        sheet = archmybook['Свод']

        for rowa in range(2, sheet.max_row):
            zip_pn = str(sheet.cell(row=rowa, column=2).value)
            app_comment = sheet.cell(row=rowa, column=12).value
            stock = sheet.cell(row=rowa, column=8).value

            filter_key = self.filterkey(key=zip_pn).upper()

            if filter_key not in archdict:
                item = {'ZIP': zip_pn,
                        'Appoint': f'{app_comment}',
                        'MainKey': f'{zip_pn}',
                        'Stock': f'{stock}'}

                archdict[filter_key] = item

        return archdict

    def orderstatus(self, num_order, orderbook) -> bool:
        """search by status."""

        sheet = orderbook['Sheet1']
        for rowa in range(1, sheet.max_row):
            st_order = sheet.cell(row=rowa, column=8).value
            order = sheet.cell(row=rowa, column=1).value

            if num_order == order:

                if st_order is not None:
                    if st_order == 'отправлено':
                        return True

                return False

    def categoryarch(self, key, categorybook):
        """search by first letters in categories."""

        noneresult = {'Category': 'None',
                      'ServicePrice': 6001,
                      'Hours': 4}

        sheet = categorybook['Ценник']
        for rowa in range(3, sheet.max_row):
            category_text = str(sheet.cell(row=rowa, column=1).value)
            category = str(sheet.cell(row=rowa, column=2).value.upper())

            f_text = self.filterkey(key=category_text)
            if f_text in key:
                return {'Category': category}


        return noneresult

    def archivecreation(self, archivebook, orderbook, category):
        """archive creation."""

        archivedict: dict = {}
        sheet = archivebook['Лист1']
        for rowas in reversed(range(1, sheet.max_row)):

            main_pn = str(sheet.cell(row=rowas, column=3).value)
            zip_pn = sheet.cell(row=rowas, column=13).value
            service_comment = sheet.cell(row=rowas, column=14).value
            appoint = sheet.cell(row=rowas, column=15).value
            price_zip = sheet.cell(row=rowas, column=12).value
            num_order = sheet.cell(row=rowas, column=24).value
            qty = sheet.cell(row=rowas, column=5).value

            atribute = {'ZIP': f'{zip_pn}',
                        'ServiceComment': f'{service_comment}',
                        'Appoint': f'{appoint}',
                        'PriceCostZIP': price_zip,
                        'NumOrder': f'{num_order}',
                        'MainKey': f'{main_pn}',
                        'QTY': qty,
                        'NumberStr': rowas}

            order_book = self.orderstatus(num_order=num_order, 
                                          orderbook=orderbook)

            filter_key = self.filterkey(key=main_pn).upper()

            if filter_key not in archivedict:

                archivedict[filter_key] = {'LastNumOrder': num_order}
                if num_order is not None:

                    if order_book is True:
                        if zip_pn is not None:

                            atribute.update(archivedict[filter_key])
                            archivedict[filter_key] = atribute
                            continue
            else:
                try:
                    archivedict[filter_key]['QTY'] += qty
                except:
                    continue

        return archivedict

    def categorydescription(self, param: str, book, category_book):
        """Seacrh category description."""

        local_item: dict = {}
        main_description = self.filterkey(key=param[2]).upper()
        sheet = book['Исключения']
        for rowa in range(2, sheet.max_row):

            category = str(sheet.cell(row=rowa,
                                      column=3).value).upper()

            collision = str(self.filterkey(key=sheet.cell(row=rowa,
                                           column=1).value)).upper()

            if collision in main_description:
                local_item = {'Category': category}
                l_i = self.categoryfinal(param=local_item,
                                         categoryybook=category_book)

                local_item.update(l_i)

                return local_item

        return False

    def searchebay(self, key):
        """Search Ebay."""

        result_none: dict = {'URL': 'Нет подходящих вариантов',
                             'PriceEbay': '',
                             'Shipping': '',
                             'ShippingCost': ''
                            }

        dict_ebay: dict = {}
        min_price = 1000000000
        payload = {'keywords': f'{key}',
                  'itemFilter': [{'name': 'LocatedIn',
                                 'value': 'WorldWide'}]
                  }

        try:
            api = Finding(appid=self.api_key_ebay, 
                          config_file=None, 
                          siteid='EBAY-US')

            response = api.execute('findItemsAdvanced', payload)

            if response.reply.searchResult._count == '0':
                return result_none

            for item in response.reply.searchResult.item:
                title = item.title.upper().split()
                if len(title) < 4:
                    title = self.listadd(title=title)

                if key in title:
                    price_item = float(item.sellingStatus.currentPrice.value)

                    min_price = price_item
                    delivery_price = price_item / 2
                    dict_ebay = {'URL': item.viewItemURL,
                                 'PriceEbay': price_item,
                                 'ShippingCost': delivery_price,
                                 'Shipping': '1/2PriceItem'}
                    break

            if dict_ebay == {}:
                return result_none

            return dict_ebay

        except ConnectionError as error:
            return 'ConnectionError'

    def nowritercatebay(self, item: dict, key):
        """exclusion categories."""

        search_ebay = self.searchebay(key=key)
        local_item: dict = {}
        try:
            category = item['Category']
            if category != 'LIC-1' and category != 'SOFT-1' and category != 'MSCL':
                local_item.update(search_ebay)
        except:

            local_item.update(search_ebay)

        return local_item

    def categoryfinal(self, param, categoryybook):
        """final category attributes."""

        noneresult = {'Category': 'None',
                      'ServicePrice': 6001,
                      'Hours': 4}

        sheet = categoryybook['Категории']
        for rowa in range(2, sheet.max_row):

            category = sheet.cell(row=rowa, column=1).value.upper()
            hours = sheet.cell(row=rowa, column=4).value
            service_price = sheet.cell(row=rowa, column=5).value
            
            
            if param['Category'] == category:

                return {'ServicePrice': service_price,
                        'Hours': hours,
                        'Category': category}

        return noneresult

    def archivefinding(self, local_key, 
                       archive_dict, category_book, main_key):
        """Filling with attributes."""

        local_item = {}

        try:
                search_cat_arch = self.categoryfinal(param=self.categoryarch(key=local_key, 
                                                     categorybook=category_book), 
                                                     categoryybook=category_book
                                                     )
                local_item.update(search_cat_arch)
        except:
                pass

        return local_item

    def writing(self, item_dict: dict, ws, 
                appoint: str, book: str, 
                event: str, l_count: int, 
                localKey, mainkey,
                list_item: list):
        """Writer."""

        #other_info
        stroka = 'архив: '
        stroka_pn = 'PN совпал без букв! '
        l_count += 1

        #book_ws
        ws_point, ws_archive = ws

        customer = list_item[l_count - 2][4]
        vendor = list_item[l_count - 2][3]
        description = list_item[l_count - 2][2]
        amaunt = list_item[l_count - 2][1]
        part = list_item[l_count - 2][0]

        archive_a = f'=IF(Расчет!A{l_count}="","",Расчет!A{l_count})'
        archive_b = f'=IF(Расчет!B{l_count}="","",Расчет!B{l_count})'
        archive_c = f'=IF(Расчет!C{l_count}="","",Расчет!C{l_count})'
        archive_d = f'=IF(Расчет!D{l_count}="","",Расчет!D{l_count})'
        archive_e = f'=IF(Расчет!E{l_count}="","",Расчет!E{l_count})'
        archive_k = f'=IF(Расчет!F{l_count}="","",Расчет!F{l_count})'
        archive_l = f'=IF(Расчет!G{l_count}="","",Расчет!G{l_count})'
        archive_m = f'=IF(Расчет!H{l_count}="","",Расчет!H{l_count})'
        archive_n = f'=IF(Расчет!I{l_count}="","",Расчет!I{l_count})'
        archive_o = f'=IF(Расчет!J{l_count}="","",Расчет!J{l_count})'
        archive_p = f'=IF(Расчет!P{l_count}="","",Расчет!P{l_count})'
        archive_s = f'=IF(Расчет!W{l_count}="","",Расчет!W{l_count})'
        archive_t = f'=IF(Расчет!U{l_count}="","",Расчет!U{l_count})'
        archive_u = f'=IF(Расчет!V{l_count}="","",Расчет!V{l_count})'
        archive_v = f'=IF(Расчет!O{l_count}="","",Расчет!O{l_count})'
        archive_w = f'=IF(Расчет!N{l_count}="",IF(Расчет!X{l_count}="","",Расчет!X{l_count}),Расчет!N{l_count})'

        q_form = f'=IF(R{l_count}="","",R{l_count}/2)'
        t_form = f'=IF(S{l_count}="","",L{l_count}*E{l_count}*0.1)'
        s_form = f'=IF(R{l_count}="","",(R{l_count}*2+Q{l_count})*1.15)'
        l_form = f'=IF(R{l_count}="","",R{l_count}*2+Q{l_count})'

        thin_border = Border(left=Side(style='thin'), 
                             right=Side(style='thin'), 
                             top=Side(style='thin'), 
                             bottom=Side(style='thin'))

        #WriteParamArchive
        a_archive = ws_archive.cell(row=l_count, column=1)
        a_archive.border = thin_border

        b_archive = ws_archive.cell(row=l_count, column=2)
        b_archive.border = thin_border

        c_archive = ws_archive.cell(row=l_count, column=3)
        c_archive.border = thin_border

        d_archive = ws_archive.cell(row=l_count, column=4)
        d_archive.border = thin_border     

        e_archive = ws_archive.cell(row=l_count, column=5)
        e_archive.border = thin_border

        f_archive = ws_archive.cell(row=l_count, column=6)
        f_archive.border = thin_border

        g_archive = ws_archive.cell(row=l_count, column=7)
        g_archive.border = thin_border

        h_archive = ws_archive.cell(row=l_count, column=8)
        h_archive.border = thin_border

        i_archive = ws_archive.cell(row=l_count, column=9)
        i_archive.border = thin_border

        j_archive = ws_archive.cell(row=l_count, column=10)
        j_archive.border = thin_border

        k_archive = ws_archive.cell(row=l_count, column=11)
        k_archive.border = thin_border

        l_archive = ws_archive.cell(row=l_count, column=12)
        l_archive.border = thin_border

        m_archive = ws_archive.cell(row=l_count, column=13)
        m_archive.border = thin_border

        n_archive = ws_archive.cell(row=l_count, column=14)
        n_archive.border = thin_border

        o_archive = ws_archive.cell(row=l_count, column=15)
        o_archive.border = thin_border

        p_archive = ws_archive.cell(row=l_count, column=16)
        p_archive.border = thin_border

        q_archive = ws_archive.cell(row=l_count, column=17)
        q_archive.border = thin_border

        r_archive = ws_archive.cell(row=l_count, column=18)
        r_archive.border = thin_border

        s_archive = ws_archive.cell(row=l_count, column=19)
        s_archive.border = thin_border

        t_archive = ws_archive.cell(row=l_count, column=20)
        t_archive.border = thin_border

        u_archive = ws_archive.cell(row=l_count, column=21)
        u_archive.border = thin_border

        v_archive = ws_archive.cell(row=l_count, column=22)
        v_archive.border = thin_border

        w_archive = ws_archive.cell(row=l_count, column=23)
        w_archive.border = thin_border

        #WriteParamPoint
        a = ws_point.cell(row=l_count, column=1)
        a.border = thin_border

        b = ws_point.cell(row=l_count, column=2)
        b.border = thin_border
        
        c = ws_point.cell(row=l_count, column=3)
        c.border = thin_border

        d = ws_point.cell(row=l_count, column=4)
        d.border = thin_border

        e = ws_point.cell(row=l_count, column=5)
        e.border = thin_border

        f = ws_point.cell(row=l_count, column=6)
        f.border = thin_border

        g = ws_point.cell(row=l_count, column=7)
        g.border = thin_border

        h = ws_point.cell(row=l_count, column=8)
        h.border = thin_border

        i = ws_point.cell(row=l_count, column=9)
        i.border = thin_border

        j = ws_point.cell(row=l_count, column=10)
        j.border = thin_border

        k = ws_point.cell(row=l_count, column=11)
        k.border = thin_border

        l = ws_point.cell(row=l_count, column=12)
        l.border = thin_border

        m = ws_point.cell(row=l_count, column=13)
        m.border = thin_border

        n = ws_point.cell(row=l_count, column=14)
        n.border = thin_border

        o = ws_point.cell(row=l_count, column=15)
        o.border = thin_border

        p = ws_point.cell(row=l_count, column=16)
        p.border = thin_border

        q = ws_point.cell(row=l_count, column=17)
        q.border = thin_border

        r = ws_point.cell(row=l_count, column=18)
        r.border = thin_border

        s = ws_point.cell(row=l_count, column=19)
        s.border = thin_border

        t = ws_point.cell(row=l_count, column=20)
        t.border = thin_border

        u = ws_point.cell(row=l_count, column=21)
        u.border = thin_border

        v = ws_point.cell(row=l_count, column=22)
        v.border = thin_border

        w = ws_point.cell(row=l_count, column=23)
        w.border = thin_border

        a.value = customer
        b.value = vendor
        c.value = part
        d.value = description
        e.value = amaunt

        a_archive.value = archive_a
        b_archive.value = archive_b
        c_archive.value = archive_c
        d_archive.value = archive_d
        e_archive.value = archive_e
        k_archive.value = archive_k
        l_archive.value = archive_l
        m_archive.value = archive_m
        n_archive.value = archive_n
        o_archive.value = archive_o
        p_archive.value = archive_p
        s_archive.value = archive_s
        t_archive.value = archive_t
        u_archive.value = archive_u
        v_archive.value = archive_v
        w_archive.value = archive_w

        if 'ServicePrice' in item_dict:
            if mainkey[1] != 0 and mainkey[1] is not None:
                amount = mainkey[1]
                type_amount = type(amount)
                my_item = item_dict['ServicePrice']
                type_me = type(my_item)
                if type_me is int or type_me is float:
                    if type_amount is int or type_amount is float:
                        if my_item != 0 and amount != 0:
                            f.value = my_item * amount
                        else:
                            f.value = my_item
                    else:
                        f.value = my_item

        if 'PriceCostZIP' in item_dict:
            my_item = item_dict['PriceCostZIP']
            type_me = type(my_item)
            if type_me is int or type_me is float:
                g.value = my_item

        if book == 'WriteZIP':
            if 'ZIP' in item_dict:
                my_item = item_dict['ZIP']
                if my_item != "None" and my_item != "-":
                    if appoint == 'Archive':
                        if event == 'Collision':
                            h.value = str(localKey) + ': ' + f'{my_item}'
                        else:
                            h.value = f'{my_item}'
                    elif event == 'Collision':
                        h.value = str(localKey) + ': ' + f'{my_item}'
                    else:
                        h.value = f'{my_item}'

        if 'ServiceComment' in item_dict:
            my_item = item_dict['ServiceComment']
            if my_item != "None":
                if appoint == 'Archive':
                    if event == 'Collision':
                        i.value = stroka_pn + f': {localKey} :' + f'{my_item}'
                    else:
                        i.value = f'{my_item}'

                elif event == 'Collision':
                    i.value = stroka_pn + f': {localKey} :' + f'{my_item}'
                else:
                    i.value = f'{my_item}'

        if 'ServiceComment' not in item_dict.keys():
            if event == 'Collision':
                i.value = f'{stroka_pn}' + f': {localKey} :'

        if 'Appoint' in item_dict:
            my_item = item_dict['Appoint']
            type_me = type(my_item)
            if type_me is str:
                if my_item != "None" and my_item != "-":
                    if appoint == 'Archive':
                        j.value = f'{my_item}'
                    else:
                        j.value = f'{my_item}'

        if 'NumOrder' in item_dict.keys():
            m.value = item_dict['NumOrder']

        if 'QTY' in item_dict:
            u.value = item_dict['QTY']

        if 'NumberStr' in item_dict:
            v.value = item_dict['NumberStr']

        if 'LastNumOrder' in item_dict:
            w.value = item_dict['LastNumOrder']

        if 'Stock' in item_dict:
            if item_dict['Stock'] is not None:
                p.value = item_dict['Stock']

        if 'Category' in item_dict:
            n.value = item_dict['Category']
            if mainkey[1] != 0 and mainkey[1] is not None:
                hours = item_dict['Hours']
                amount = mainkey[1]
                type_amount = type(amount)
                if type_amount is int or type_amount is float:
                    if hours != 0 and amount != 0:
                        o.value = hours * amount
                    else:
                        o.value = hours
                else:
                    o.value = hours

        if 'PriceEbay' in item_dict:
            k.value = item_dict['URL']
            if type(item_dict['ShippingCost']) is not str:
                price_ebay = item_dict['PriceEbay']
                shipping_cost = item_dict['ShippingCost']
                r.value = int(math.ceil(price_ebay))
            else:
                r.value = 0

        q.value = q_form
        l.value = l_form
        s.value = s_form
        t.value = t_form

        q.number_format = '0'
        l.number_format = '0'
        s.number_format = '0'
        t.number_format = '0'


    def collision(self, key, vendor):
        """Collision."""

        l_list = []
        key_main_filter = self.filterkey(key=key)
        l_list.append(key_main_filter)

        if vendor == 'CISCO' and 'R-' in key:
            key_one_filter = self.filterkey(key=key.replace('R-', '-'))
            l_list.append(key_one_filter)
        try:
            if '24' in key:
                number = '0123456789'
                for i in range(0, len(key)):

                    if key[i] == '2' and i != len(key):
                        if key[i + 1] == '4' and i + 2 <= len(key):

                            if key[i + 2] not in number:
                                if i - 1 >= 0 and key[i - 1] not in number:
                                    key_one_filter = self.filterkey(key=key.replace('24','48'))
                                    l_list.append(key_one_filter)

                                elif i - 1 < 0:
                                    key_one_filter = self.filterkey(key=key.replace('24', '48'))
                                    l_list.append(key_one_filter)

                        elif key[i + 1] == '4' and i + 2 > len(key):
                            if i - 1 >= 0:
                                if key[i - 1] not in number:
                                    key_one_filter = self.filterkey(key=key.replace('24', '48'))
                                    l_list.append(key_one_filter)
        except:
            pass

        if 'K7' in key:
            key_one_filter = self.filterkey(key=key.replace('K7', 'K8'))
            key_two_filter = self.filterkey(key=key.replace('K7', 'K9'))
            l_list.append(key_one_filter)
            l_list.append(key_two_filter)

        if 'K8' in key:
            key_one_filter = self.filterkey(key=key.replace('K8', 'K9'))
            key_two_filter = self.filterkey(key=key.replace('K8', 'K7'))
            l_list.append(key_one_filter)
            l_list.append(key_two_filter)

        if 'K9' in key:
            key_one_filter = self.filterkey(key=key.replace('K9', 'K7'))
            key_two_filter = self.filterkey(key=key.replace('K9', 'K8'))
            l_list.append(key_one_filter)
            l_list.append(key_two_filter)           

        my_tuple = (l_list, len(l_list))
        return my_tuple

    def listadd(self, title: list):
        """Optional function."""

        if len(title) == 1:
            title.append('1')
            title.append('2')
            title.append('6')
        if len(title) == 2:
            title.append('3')
            title.append('0')
        if len(title) == 3:
            title.append('4')

        return title

    def filterkey(self, key):
        """FilterKeys func."""

        key = key.replace(' ', '')
        getvals = list([val for val in key if val.isalpha() or val.isnumeric()])
        result = "".join(getvals)
        return result

