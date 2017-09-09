import datetime
import os

import requests
import xlsxwriter
from multiprocessing.dummy import Pool as ThreadPool

session = requests.session()
pool = ThreadPool(10)
headers = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64; rv:52.0) Gecko/20100101 Firefox/52.0",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.5",
    "Connection": "keep-alive",
    "Host": "www.hollandamerica.com",
    "Adrum": "isAjax:true",
    "Refer": "http://www.hollandamerica.com/find-cruise-vacations/FindCruises?showSoldOut=true&page=0&cfVer=1"
}

urls = []
voyage_ids = []
voyages = []
to_write = []
session.headers.update(headers)
url = "http://www.hollandamerica.com/assets/jsonroot/hal/indexes/search/v1_1/index.json"
page = session.get(url)
cruise_results = page.json()
for k, v in cruise_results["index"].items():
    if k == "ship":
        for each in v:
            for line in v[each]:
                split = line.split("-")
                if line not in urls:
                    urls.append(split[1])


def convert_date(not_formatted):
    splitter = not_formatted.split("-")
    day = splitter[2]
    month = splitter[1]
    year = splitter[0]
    final_date = '%s/%s/%s' % (month, day, year)
    return final_date


def get_destination(param):
    if param == 'A':
        return ['Alaska', 'A']
    elif param == 'O':
        return ['Asia & Pacific', 'O']
    elif param == 'P':
        return ['Australia/New Zealand & S.Pacific', 'P']
    elif param == 'B':
        return ['Bermuda', 'B']
    elif param == 'N':
        return ['Canada/New England', 'N']
    elif param == 'C':
        return ['Caribbean', 'C']
    elif param == 'E':
        return ['Europe', 'E']
    elif param == 'W':
        return ['Grand Voyages', 'W']
    elif param == 'H':
        return ['Hawaii & Tahiti', 'H']
    elif param == 'X':
        return ['Holiday', 'X']
    elif param == 'M':
        return ['Mexico', 'M']
    elif param == 'L':
        return ['Pacific Northwest & Pacific Coast', 'L']
    elif param == 'T':
        return ['Panama Canal', 'T']
    elif param == 'S':
        return ['South America & Antarctica', 'S']
    elif param == 'G':
        return ['Gonna know', 'G']
    elif param == 'U':
        return ['Gonna know', 'U']
    elif param == 'R':
        return ['Gonna know', 'R']
    elif param == 'Z':
        return ['Gonna know', 'Z']
    elif param == 'F':
        return ['Gonna know', 'F']
    elif param == 'I':
        return ['Gonna know', 'I']
    elif param == 'Y':
        return ['Gonna know', 'Y']
    elif param == 'V':
        return ['Gonna know', 'V']


single = []
for u in list(set(urls)):
    url = "http://www.hollandamerica.com/assets/jsonroot/hal/itineraries/v1_0/USD/" + u + ".json"
    single.append(url)
single = list(set(single))


def get_vessel_id(name):
    if name == "Amsterdam":
        return "108"
    if name == "Eurodam":
        return "580"
    if name == "Koningsdam":
        return "926"
    if name == "Maasdam":
        return "110"
    if name == "Nieuw Amsterdam":
        return "719"
    if name == "Noordam":
        return "496"
    if name == "Oosterdam":
        return "410"
    if name == "Prinsendam":
        return "407"
    if name == "Rotterdam":
        return "113"
    if name == "Veendam":
        return "118"
    if name == "Volendam":
        return "119"
    if name == "Westerdam":
        return "434"
    if name == "Zaandam":
        return "121"
    if name == "Zuiderdam":
        return "409"


counter = 0
codes = []


def parse(ur):
    result_page = session.get(ur)
    results = result_page.json()
    if ur == 'http://www.hollandamerica.com/assets/jsonroot/hal/itineraries/v1_0/USD/A8G07X.json':
        for sailing in results['voyages']:
            print(sailing)
    for cr in results['voyages']:
        # if line['voyageId'] in codes:
        #     continue
        # else:
        #     codes.append(line['voyageId'])
        if 'ITINERARY' in cr['itineraryType']:
            pass
        elif 'TOUR' in cr['itineraryType']:
            continue
        print("Processing", ur)
        brochure_name = results['description']
        cruise_line_name = "Holland America"
        interior_bucket_price = "N/A"
        balcony_bucket_price = "N/A"
        ocean_view_bucket_price = "N/A"
        suite_bucket_price = ""
        signature = ''
        neptune = ''
        obstructed = ''
        cruise_id = "8"
        itinerary_id = cr['voyageId']
        number_of_nights = (cr['duration'])
        destination = get_destination(cr['direction'])
        destination_name = destination[0]
        destination_code = destination[1]
        sail_date = convert_date(cr['dateDepart'])
        return_date = convert_date(cr['dateArrive'])
        vessel_name = cr['ship']['displayName'].replace('ms ', '')
        vessel_id = get_vessel_id(vessel_name)
        for room in cr['stateRooms']:
            if room['id'] == 'Interior':
                if room["priceBlocks"][0]['currencyCode'] == "SOLD_OUT":
                    interior_bucket_price = 'N/A'
                else:
                    if room["priceBlocks"][0]['prices'][0]['campaignId'] == "HALBEST":
                        interior_bucket_price = room["priceBlocks"][0]['prices'][0]['fare'].split('.')[0]
            elif room['id'] == "Ocean-view":
                if room["priceBlocks"][0]['currencyCode'] == "SOLD_OUT":
                    ocean_view_bucket_price = 'N/A'
                else:
                    if room["priceBlocks"][0]['prices'][0]['campaignId'] == "HALBEST":
                        ocean_view_bucket_price = room["priceBlocks"][0]['prices'][0]['fare'].split('.')[0]
            elif room['id'] == 'Verandah':
                if room["priceBlocks"][0]['currencyCode'] == "SOLD_OUT":
                    balcony_bucket_price = 'N/A'
                else:
                    if room["priceBlocks"][0]['prices'][0]['campaignId'] == "HALBEST":
                        balcony_bucket_price = room["priceBlocks"][0]['prices'][0]['fare'].split('.')[0]
            elif room['id'] == 'Vista Suite':
                if room["priceBlocks"][0]['currencyCode'] == "SOLD_OUT":
                    suite_bucket_price = 'N/A'
                else:
                    if room["priceBlocks"][0]['prices'][0]['campaignId'] == "HALBEST":
                        suite_bucket_price = room["priceBlocks"][0]['prices'][0]['fare'].split('.')[0]
            elif room['id'] == 'Signature Suite':
                if room["priceBlocks"][0]['currencyCode'] == "SOLD_OUT":
                    signature = 'N/A'
                else:
                    if room["priceBlocks"][0]['prices'][0]['campaignId'] == "HALBEST":
                        signature = room["priceBlocks"][0]['prices'][0]['fare'].split('.')[0]
            elif room['id'] == 'Neptune Suite':
                if room["priceBlocks"][0]['currencyCode'] == "SOLD_OUT":
                    neptune = 'N/A'
                else:
                    if room["priceBlocks"][0]['prices'][0]['campaignId'] == "HALBEST":
                        neptune = room["priceBlocks"][0]['prices'][0]['fare'].split('.')[0]
            elif room['id'] == 'Obstructed Verandah':
                if room["priceBlocks"][0]['currencyCode'] == "SOLD_OUT":
                    obstructed = 'N/A'
                else:
                    if room["priceBlocks"][0]['prices'][0]['campaignId'] == "HALBEST":
                        obstructed = room["priceBlocks"][0]['prices'][0]['fare'].split('.')[0]
        if vessel_name == "Amsterdam" or vessel_name == 'Maasdam' or vessel_name == 'Prinsendam' or vessel_name == 'Rotterdam' or vessel_name == 'Veendam' or vessel_name == 'Volendam' or vessel_name == 'Zaandam':
            if suite_bucket_price != '':
                balcony_bucket_price = suite_bucket_price
            else:
                balcony_bucket_price = 'N/A'
            if signature != '':
                suite_bucket_price = signature
            else:
                if neptune != '':
                    suite_bucket_price = neptune
                else:
                    suite_bucket_price = "N/A"
        elif vessel_name == "Eurodam" or vessel_name == 'Nieuw Amsterdam' or vessel_name == 'Noordam' or vessel_name == 'Oosterdam' or vessel_name == 'Prinsendam' or vessel_name == 'Westerdam' or vessel_name == 'Zuiderdam':
            if signature == '':
                suite_bucket_price = neptune
            else:
                suite_bucket_price = signature
        elif vessel_name == 'Koningsdam':
            if balcony_bucket_price == '' and obstructed == '':
                if suite_bucket_price == '':
                    balcony_bucket_price = 'N/A'
                else:
                    balcony_bucket_price = suite_bucket_price
            else:
                if obstructed == '':
                    pass
                else:
                    balcony_bucket_price = obstructed

            if suite_bucket_price == "":
                if signature == '':
                    if neptune == '':
                        suite_bucket_price = "N/A"
                    else:
                        suite_bucket_price = neptune
                else:
                    suite_bucket_price = signature
            else:
                pass

        temp = [destination_code, destination_name, vessel_id, vessel_name, cruise_id, cruise_line_name, itinerary_id,
                brochure_name, number_of_nights, sail_date, return_date, str(interior_bucket_price),
                str(ocean_view_bucket_price), str(balcony_bucket_price), str(suite_bucket_price)]
        if temp in to_write:
            pass
        else:
            to_write.append(temp)


pool.map(parse, single)
pool.close()
pool.join()


def write_file_to_excell(data_array):
    userhome = os.path.expanduser('~')
    now = datetime.datetime.now()
    path_to_file = userhome + '/Dropbox/XLSX/For Assia to test/' + str(now.year) + '-' + str(now.month) + '-' + str(
        now.day) + '/' + str(now.year) + '-' + str(now.month) + '-' + str(now.day) + '- Holland America.xlsx'
    if not os.path.exists(userhome + '/Dropbox/XLSX/For Assia to test/' + str(now.year) + '-' + str(
            now.month) + '-' + str(now.day)):
        os.makedirs(
            userhome + '/Dropbox/XLSX/For Assia to test/' + str(now.year) + '-' + str(now.month) + '-' + str(now.day))
    workbook = xlsxwriter.Workbook(path_to_file)

    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})
    worksheet.set_column("A:A", 15)
    worksheet.set_column("B:B", 25)
    worksheet.set_column("C:C", 10)
    worksheet.set_column("D:D", 25)
    worksheet.set_column("E:E", 20)
    worksheet.set_column("F:F", 30)
    worksheet.set_column("G:G", 20)
    worksheet.set_column("H:H", 50)
    worksheet.set_column("I:I", 20)
    worksheet.set_column("J:J", 20)
    worksheet.set_column("K:K", 20)
    worksheet.set_column("L:L", 20)
    worksheet.set_column("M:M", 25)
    worksheet.set_column("N:N", 20)
    worksheet.set_column("O:O", 20)
    worksheet.write('A1', 'DestinationCode', bold)
    worksheet.write('B1', 'DestinationName', bold)
    worksheet.write('C1', 'VesselID', bold)
    worksheet.write('D1', 'VesselName', bold)
    worksheet.write('E1', 'CruiseID', bold)
    worksheet.write('F1', 'CruiseLineName', bold)
    worksheet.write('G1', 'ItineraryID', bold)
    worksheet.write('H1', 'BrochureName', bold)
    worksheet.write('I1', 'NumberOfNights', bold)
    worksheet.write('J1', 'SailDate', bold)
    worksheet.write('K1', 'ReturnDate', bold)
    worksheet.write('L1', 'InteriorBucketPrice', bold)
    worksheet.write('M1', 'OceanViewBucketPrice', bold)
    worksheet.write('N1', 'BalconyBucketPrice', bold)
    worksheet.write('O1', 'SuiteBucketPrice', bold)
    row_count = 1
    money_format = workbook.add_format({'bold': True})
    ordinary_number = workbook.add_format({"num_format": '#,##0'})
    date_format = workbook.add_format({'num_format': 'm d yyyy'})
    centered = workbook.add_format({'bold': True})
    money_format.set_align("center")
    money_format.set_bold(True)
    date_format.set_bold(True)
    centered.set_bold(True)
    ordinary_number.set_bold(True)
    ordinary_number.set_align("center")
    date_format.set_align("center")
    centered.set_align("center")
    for ship_entry in data_array:
        column_count = 0
        for en in ship_entry:
            if column_count == 0:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 1:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 2:
                try:
                    worksheet.write_string(row_count, column_count, en, centered)
                except TypeError:
                    worksheet.write_string(row_count, column_count, " ", centered)
            if column_count == 3:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 4:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 5:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 6:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 7:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 8:
                try:
                    worksheet.write_number(row_count, column_count, en, ordinary_number)
                except TypeError:
                    worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 9:
                try:
                    date_time = datetime.datetime.strptime(str(en), "%m/%d/%Y")
                    worksheet.write_datetime(row_count, column_count, date_time, money_format)
                except TypeError:
                    worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 10:
                try:
                    date_time = datetime.datetime.strptime(str(en), "%m/%d/%Y")
                    worksheet.write_datetime(row_count, column_count, date_time, money_format)
                except TypeError:
                    worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 11:
                try:
                    worksheet.write_number(row_count, column_count, int(en), money_format)
                except ValueError:
                    worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 12:
                try:
                    worksheet.write_number(row_count, column_count, int(en), money_format)
                except ValueError:
                    worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 13:
                try:
                    worksheet.write_number(row_count, column_count, int(en), money_format)
                except ValueError:
                    worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 14:
                try:
                    worksheet.write_number(row_count, column_count, int(en), money_format)
                except ValueError:
                    worksheet.write_string(row_count, column_count, en, centered)
            column_count += 1
        row_count += 1
    workbook.close()
    pass


write_file_to_excell(to_write)
print('Itineraries:', len(single))
print('Voyages:', len(to_write))
