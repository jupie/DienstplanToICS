from datetime import datetime
from icalendar import Calendar, Event
import easygui
from dataclasses import dataclass
from bs4 import BeautifulSoup
import bs4.element
import csv

month = -1
year = -1


@dataclass
class Mitarbeiter:
    name: str
    dienstestring: []
    dienste: []


class Dienst:
    name: str =''
    beschreibung: str = ''
    coworker: str = ''
    vehicle: str =''
    start: datetime
    end: datetime

    def __init__(self, name, coworker, day):
        day = day + 1
        self.name = name
        self.coworker = coworker

        if 'Frei' in name:
            self.start = datetime(year, month, day, 00, 00)
            self.end = datetime(year, month, day, 23, 59)
            self.vehicle = 'Frei'
            self.coworker = ''
        else:
            self.parse_dienst(day)

    def parse_dienst(self, day):

        global dienste_array
        dienst_row = []
        for row in dienste_array:
            if row[0].startswith(self.name):
                dienst_row = row
                break
        self.vehicle = dienst_row[1]
        self.beschreibung = dienst_row[2]
        beginhour = int(dienst_row[3].split(':')[0])
        self.start = datetime(year, month, day, beginhour, int(dienst_row[3].split(':')[1]))
        endhour = int(dienst_row[4].split(':')[0])
        if beginhour >= endhour:
            self.end = datetime(year, month, day + 1, endhour,int(dienst_row[4].split(':')[1]) )
        else:
            self.end = datetime(year, month, day, endhour, int(dienst_row[4].split(':')[1]))

    def get_event(self):
        event = Event()
        if self.vehicle and self.coworker:
            event.add('summary', self.vehicle + ' mit ' + self.coworker)
        else:
            event.add('summary',self.name)


        event.add('dtstart', self.start)
        event.add('dtend', self.end)
        return event


def parse_people(table: bs4.Tag):
    people = []
    for tag in table.find_all('th'):
        people.append(tag.string)

    return people[1:]


def parse_dienste(table: bs4.Tag, people: []):
    rows: [] = table.find_all('tr')
    mitarbeiter: [] = []
    for idx, row in enumerate(rows[1:]):
        dienste = []
        for day in row.find_all('td'):
            daystring = day.string
            if daystring:
                dienste.append(daystring)
            else:
                dienste.append('Frei')

        mitarbeiter.append(Mitarbeiter(people[idx], dienste, []))
    return mitarbeiter


def select_mtarbeiter(mitarbeiter):
    names = [ma.name for ma in mitarbeiter]
    return easygui.choicebox('Wählen Sie einen Mitarbeiter', '', names)


def find_coworker(name, day, dienststring, mitarbeiterarray):
    for mitarbeiter in mitarbeiterarray:
        if not mitarbeiter.name.startswith(name) and mitarbeiter.dienstestring[day].startswith(dienststring):
            return mitarbeiter.name


def parse_mitarbeiter_coworker(ma, mitarbeiterarray):
    dienste = []
    mitarbeiter = None
    for mitarbeitercotainer in mitarbeiterarray:
        if mitarbeitercotainer.name.startswith(ma):
            mitarbeiter = mitarbeitercotainer
    if not mitarbeiter:
        raise Exception('Mitarbeiter nicht vorhanden')

    for day, dienststring in enumerate(mitarbeiter.dienstestring):
        if dienststring.startswith('Frei'):
            dienste.append(Dienst(dienststring, '', day))
            continue
        coworker = find_coworker(mitarbeiter.name, day, dienststring, mitarbeiterarray)
        dienste.append(Dienst(dienststring, coworker, day))
    mitarbeiter.dienste = dienste
    return mitarbeiter


def set_year_and_month(soup: bs4.Tag):
    global month
    global year
    monthstring = soup.find("option", attrs={"selected": "selected"}).string
    monthbyname = monthstring.split(' ')[0]
    months = {"Januar": 1, "Februar": 2, "März": 3, "April": 4, "Mai": 5, "Juni": 6, "Juli": 7, "August": 8,
              "September": 9,
              "Oktober": 10, "November": 11, "Dezember": 12}
    month = months[monthbyname]
    year = int(monthstring.split(' ')[1])


def parse_dienstplanexport(file):
    with open(file, encoding='utf-8') as export:
        soup = BeautifulSoup(export, 'html.parser')
        set_year_and_month(soup)
        tables = soup.find_all('table')
        people = parse_people(tables[0])
        mitarbeiterarray = parse_dienste(tables[1], people)
        ma = select_mtarbeiter(mitarbeiterarray)
        mitarbeiter = parse_mitarbeiter_coworker(ma, mitarbeiterarray)
        build_ical(mitarbeiter)


def build_ical(mitarbeiter):
    cal = Calendar()
    cal.add('prodid', '-//My calendar product//mxm.dk//')
    cal.add('version', '2.0')
    for dienst in mitarbeiter.dienste:
        event = dienst.get_event()
        cal.add_component(event)
    saveplace = easygui.filesavebox('Datei Speichern unter',default='%HOMEPATH%/out.ics',filetypes='\\*.ics')
    with open(saveplace, 'wb') as out:
        out.write(cal.to_ical())


if __name__ == '__main__':
    easygui.msgbox('#################################\n'
                   '######Dienstplan Parser##########\n'
                   '#######by Jannik Upmann##########\n'
                   '#######jannik-upmann.de##########\n'
                   '#################################\n'
                   '\n'
                   'Bitte Wählen Sie Ihren Dienstplan export'
                   '')
    result = easygui.fileopenbox('Bitte Wählen Sie Ihren Dienstplan export')
    with open('./dienste.csv') as csv_file:
        reader = csv.reader(csv_file)
        dienste_array = [line for line in reader]

    parse_dienstplanexport(result)
