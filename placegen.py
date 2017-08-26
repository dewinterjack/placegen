#July 2017
#Developer: Jack Dewinter


import urllib.request, json, xlwt, time
from tkinter import *
import tkinter
from collections import defaultdict

#Nearby
'''
def GoogPlac(lat,lng,type,key,token):
  #making the url
  AUTH_KEY = key
  LOCATION = str(lat) + "," + str(lng)
  TYPES = type
  MyUrl = ('https://maps.googleapis.com/maps/api/place/nearbysearch/json'
           '?location=%s'
           '&rankby=distance'
           '&types=%s'
           '&sensor=false&key=%s') % (LOCATION, TYPES, AUTH_KEY)
  #grabbing the JSON result
  if token != '':
      MyUrl = MyUrl + "&pagetoken=" + token
  response = urllib.request.urlopen(MyUrl)
  jsonRaw = response.read()
  jsonData = json.loads(jsonRaw)
  return jsonData
'''
#Radar

def GoogPlac(lat,lng,type,key,token,radius):
  #making the url
  AUTH_KEY = key
  LOCATION = str(lat) + "," + str(lng)
  #LOCATION = '51.526728,-0.197135' #Custom starting location
  TYPES = type
  RADIUS = radius
  #CHANGE THIS AND THE PRINTED VARIALBE WHEN SPECIFYING
  KEYWORD = type
  MyUrl = ('https://maps.googleapis.com/maps/api/place/radarsearch/json'
           '?location=%s'
           '&radius=%s'
           '&types=%s'
           '&keyword=%s' 
           '&sensor=false&key=%s') % (LOCATION, RADIUS, TYPES, KEYWORD, AUTH_KEY)
  #grabbing the JSON result
  if token != '':
      MyUrl = MyUrl + "&pagetoken=" + token
  response = urllib.request.urlopen(MyUrl)
  jsonRaw = response.read()
  jsonData = json.loads(jsonRaw)
  return jsonData

def PlacDetails(key, id):
    AUTH_KEY = key
    PLACE = id
    MyUrl = ('https://maps.googleapis.com/maps/api/place/details/json''?placeid=%s''&key=%s') % (PLACE, AUTH_KEY)
    response = urllib.request.urlopen(MyUrl)
    jsonRaw = response.read()
    jsonData = json.loads(jsonRaw)
    return jsonData

#This is a helper to grab the Json data that I want in a list

def detailsJson(place):
    #Call Hunter API and set an email field to the result.
    x = [place['name']]
    address = ''
    postCode = 'None'
    addr = place['address_components']
    address2 = ''
    types = ''
    website = ''
    for type in place['types']:
        types = types + type + ','
    print(place['name'])
    for comp in addr:
        for type in comp['types']:
            if type == 'floor' or type == 'street_number' or type == 'route':
                address = address + comp['long_name'] + ' '
            elif type == 'postal_code':
                postCode = comp["long_name"]
            elif type == 'postal_town':
                address2 = address2 + comp['long_name'] + ','
            elif type == 'neighborhood':
                address2 = address2 + comp['long_name'] + ','

    if address == '':
        address = 'None'
    x.append(address.strip(' \t\n\r'))
    x.append(address2.strip(','))
    x.append(postCode)
    #x = [place['formatted_address']]
    if 'formatted_phone_number' in place:
        x.append(place['formatted_phone_number'])
    else:
        x.append('None')
    x.append(types.strip(','))
    if 'website' not in place:
        website = 'None'
    else:
        website = place['website']
    x.append(website)
    return x

def printDetails(sheet1,id,row,key,type):
    details = PlacDetails(key, id)
    if details['status'] == 'OK':
        place = details['result']
        x = detailsJson(place)
        #Place Name
        sheet1.write(row, 0, x[0])
        #Address
        sheet1.write(row, 1, x[1])
        #Address2
        sheet1.write(row, 2, x[2])
        # PostCode
        sheet1.write(row, 3, x[3])
        #Telephone
        sheet1.write(row, 4, x[4])
        # Types
        sheet1.write(row, 8, x[5])
        #Called
        sheet1.write(row, 10, 'FALSE')
        #Placeid
        sheet1.write(row, 11, id)
        #Keyword
        sheet1.write(row, 12, type)
        #Website
        sheet1.write(row, 13, x[6])

def printNames(sheet,search, c,key,type):
    print(search['status'])
    if search['status'] == 'OK':  # loop through results
        print(len(search))
        for place in search['results']:
            '''
            if c % 7 == 0:
                time.sleep(2)
            '''
            printDetails(sheet,place['place_id'], c,key,type)
            c = c + 1
    '''
    if 'next_page_token' in search:
        token = search['next_page_token']
        search = GoogPlac('51.544733', '-0.176414', type, key,token,radius)
        printNames(sheet,search, c,key)'''

KEY = 'AIzaSyDyf_f_6aLxBpX19MLO9jc2Ua8aIAuTG5c'
hunterKEY = '6b599787b99089db5c6bd8ccfbce71efb2c8b400'
def createSheet(KEY,type,radius,option):
    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Local " + type + "s")

    nameColumn = sheet1.col(0)
    nameColumn.width = 256 * 40

    addressColumn = sheet1.col(1)
    addressColumn.width = 256 * 30

    address2Column = sheet1.col(2)
    address2Column.width = 256 * 20

    pCodeColumn = sheet1.col(3)
    pCodeColumn.width = 256 * 10

    telColumn = sheet1.col(4)
    telColumn.width = 256 * 20

    emailColumn = sheet1.col(9)
    emailColumn.width = 256 * 30

    commentColumn = sheet1.col(7)
    commentColumn.width = 256 * 20

    typeColumn = sheet1.col(8)
    typeColumn.width = 256 * 10

    websiteColumn = sheet1.col (13)
    websiteColumn.width = 256 * 10

    #y,x
    sheet1.write(0, 0, "Place")
    sheet1.write(0, 1, "Address1")
    sheet1.write(0, 2, "Address2")
    sheet1.write(0, 3, "PostCode")
    sheet1.write(0, 4, "Tel")
    sheet1.write(0, 5, "GM")
    sheet1.write(0, 6, "DM")
    sheet1.write(0, 7, "Comments")
    sheet1.write(0, 8, "Types")
    sheet1.write(0, 9, "Email")
    sheet1.write(0, 10, "Called")
    sheet1.write(0, 11, "Placeid")
    sheet1.write(0, 12, "Keyword")
    sheet1.write(0, 13, "Website")

    c = 1
    if option == 0:
        search = GoogPlac('51.544733','-0.176414', type, KEY,'',radius) #First run no page toke needed
        printNames(sheet1, search, c, KEY, type)
    else:
        search = textSearch(KEY,option)
        print(search)
    root.destroy()
    book.save("local" + type + ".xls")


types = ["Hotel","Bar","Gym","Bakery","Bank","Beauty Salon", "Cafe","Taxi","Pharmacy","Establishment","Moving Company","Car Rental","Hospital","Mosque","Museum","Church","Dentist","Accounting","Airport","Amusement Park","Aquarium","Art Gallery","ATM","Bicycle Store","Book Store","Bowling Alley","Bus Station","Campground","Car Dealer","Car Repair","Car Wash","Casino","Cemetery","City Hall","Clothing Store","Convenience Store","Courthouse","Department Store","Doctor","Electrician","Electronics Store","Embassy","Florist","Health","Hair Care","Hardware Store","Hindu Temple","Home Goods Store","Insurance Agency","Jewelry Store","Laundry","Lawyer","Library","Liquor Store","Local Government Office","Locksmith","Meal Delivery","Meal Takeaway","Movie Rental","Movie Theater","Night Club","Painter","Park","Pet Store","Physiotherapist","Plumber","Police","Post Office","Real Estate Agency","Restaurant","Roofing Contractor","School","Shoe Store","Shopping Mall","Spa","Stadium","Synagogue","Train Station","Travel Agency","University"]
ranges = ["500m","1000m","1 mile", "2 miles", "3 miles", "4 miles", "5 miles", "10 miles"]

def typeName(type):
    return {
        'Hotel':'lodging',
        'Bar':'bar',
        'Gym':'gym',
        'Bakery':'bakery',
        'Bank':'bank',
        'Beauty Salon':'beauty_salon',
        'Cafe':'cafe',
        'Taxi':'taxi_stand',
        'Pharmacy':'pharmacy',
        'Establishment':'establishment',
        'Moving Company':'moving_company',
        'Car Rental':'car_rental',
        'Hospital':'hospital',
        'Mosque':'mosque',
        'Museum':'museum',
        'Church':'church',
        'Dentist':'dentist',
        'Accounting':'accounting',
        "Airport":'airport',
        "Amusement Park":'amusement_park',
        "Aquarium":'aquarium',
        "Art Gallery":'art_gallery',
        "ATM":'atm',
        "Bicycle Store":'bicycle_store',
        "Book Store":'book_store',
        "Bowling Alley":'bowling_alley',
        "Bus Station":'bus_station',
        "Campground":'campground',
        "Car Dealer":'car_dealer',
        "Car Repair":'car_repair',
        "Car Wash":'car_wash',
        "Casino":'casino',
        "Cemetery":'cemetery',
        "City Hall":'city_hall',
        "Clothing Store":'clothing_store',
        "Convenience Store":'convenience_store',
        "Courthouse":'courthouse',
        "Department Store":'department_store',
        "Doctor":'doctor',
        "Electrician":'electrician',
        "Electronics Store":'electronics_store',
        "Embassy":'embassy',
        "Florist":'florist',
        "Health":'health',
        "Hair Care":'hair_care',
        "Hardware Store":'hardware_store',
        "Hindu Temple":'hindu_temple',
        "Home Goods Store":'home_goods_store',
        "Insurance Agency":'insurance_agency',
        "Jewelry Store":'jewelry_store',
        "Laundry":'laundry',
        "Lawyer":'lawyer',
        "Library":'library',
        "Liquor Store":'liquor_store',
        "Local Government Office":'local_government_office',
        "Locksmith":'locksmith',
        "Meal Delivery":'meal_delivery',
        "Meal Takeaway":'meal_takeaway',
        "Movie Rental":'movie_rental',
        "Movie Theater":'movie_theater',
        "Night Club":'night_club',
        "Painter":'painter',
        "Park":'park',
        "Pet Store":'pet_store',
        "Physiotherapist":'physiotherapist',
        "Plumber":'plumber',
        "Police":'police',
        "Post Office":'post_office',
        "Real Estate Agency":'real_estate_agency',
        "Restaurant":'restaurant',
        "Roofing Contractor":'roofing_contractor',
        "School":'school',
        "Shoe Store":'shoe_store',
        "Shopping Mall":'shopping_mall',
        "Spa":'spa',
        "Stadium":'stadium',
        "Synagogue":'synagogue',
        "Train Station":'train_station',
        "Travel Agency":'travel_agency',
        "University":'university'


    }[type]
def radiusName(radius):
    return {
        '500m': 500,
        '1000m': 1000,
        '1 mile': 1600,
        '2 miles': 3200,
        '3 miles': 4800,
        '4 miles': 6400,
        '5 miles': 7600,
        '10 miles': 15200,
    }[radius]

def textSearch(key,query):
    MyUrl = 'https://maps.googleapis.com/maps/api/place/textsearch/json''?query=%s''&key=%s'%(query,key)
    response = urllib.request.urlopen(MyUrl)
    jsonRaw = response.read()
    jsonData = json.loads(jsonRaw)
    return jsonData

root = tkinter.Tk()
typeVal = StringVar(root)
typeVal.set(types[0])
radiusVal = StringVar(root)
radiusVal.set(ranges[0])
inst = Label(root,text="Select a category:")
inst.pack()
root.geometry("200x200")
root.resizable(width=False, height=False)

root.wm_title("PlaceGen")

typeOptions = OptionMenu(root,typeVal,*types)
typeOptions.pack()

rInst = Label(root,text="Select a range:")
rInst.pack()
rangeOptions = OptionMenu(root,radiusVal,*ranges)
rangeOptions.pack()


#textQuery = Button(root,text='Text Search', command= lambda: createSheet(KEY,typeName(typeVal.get()),radiusName(radiusVal.get()),query.get("1.0","end")))
#textQuery.pack()
#word = keyW.get("1.0","end")
#keyW.delete("1.0","end")
#placeToSheet = Button(root,text='PlaceID To Sheet', command= lambda: createSheet(KEY,typeName(typeVal.get()),radiusName(radiusVal.get()),1))
start = Button(root,text='Generate',command= lambda: createSheet(KEY,typeName(typeVal.get()),radiusName(radiusVal.get()),0))
start.pack()
#placeToSheet.pack()
root.mainloop()


#createSheet(KEY,'lodging',1000)