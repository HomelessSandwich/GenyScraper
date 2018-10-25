"""
====================================================================================================
|   Title:       Geny Scraper
|   Author:      Oliver Wilkins
|   Language:    Python 3.6
|   Date:        September 2018
|   To Compile:  Scrapy and xlwt will be needed to be downloaded using pip. cd to the folder in
|                which this file is located in, then use the command
|                "scrapy runspider geny_scrape.py".
|   Purpose:     Using the Scrapy framework, this program will ask the user for a date, then the
|                program will scrape all the stats from that day from www.geny.com and put it into
|                a formatted Excel file.
====================================================================================================
"""

# -*- coding: utf-8 -*-
import scrapy
from datetime import datetime
import xlwt
from tkinter import *
from tkinter import filedialog
import time
import os
import re


class Date():
    def __repr__(self):
        return f'Date: {self.date}'

    def get_date(self):
        """Asks the user for a date. Deals with error handling."""
        while True:
            try:
                self.date = input("Date (DD/MM/YYYY): ")
            except ValueError:
                print('\nThat was not was valid date!')
            else:
                break

    @property
    def date(self):
        """User's date. Needs to be in the DD/MM/YYYY format."""
        return self._date

    @date.setter
    def date(self, date):
        if self.validate_date(date):
            self._date = date
        else:
            raise ValueError('Not a valid date!')

    @property
    def day(self):
        return self.date.split('/')[0]

    @property
    def month(self):
        return self.date.split('/')[1]

    @property
    def year(self):
        return self.date.split('/')[2]

    @staticmethod
    def validate_date(date):
        """Validates a given date. Needs to be in the DD/MM/YYYY format to be excepted."""
        try:
            datetime.strptime(date, '%d/%m/%Y')
        except ValueError:
            return False
        else:
            return True


class GenyScrapeSpider(scrapy.Spider):
    name = 'geny-scrape'
    allowed_domains = ['geny.com']
    BASE_URL = 'http://www.geny.com/reunions-courses-pmu?date='

    def __init__(self):
        self.date = Date()
        self.book = xlwt.Workbook()
        self.row_index = 2  # First row of the Excel file starts on 3 (2 is zero point reference)
        self.data_style = xlwt.easyxf('font: name Arial, bold on, height 160;'
                                      'borders: top_color black, bottom_color black,'
                                      'right_color black, left_color black,'
                                      'left thin, right thin, top thin, bottom thin;'
                                      'align: horiz left;')
        self.stats = []  # For storing race stats for Excel writing
        self.rapport_data = []  # For storing money details from rapports for Excel writing

    @staticmethod
    def get_file_loc(day, month, year):
        """Allows the user to pick a folder to save a file to. Returns the file path."""
        while True:
            root = Tk()
            root.filename = filedialog.askdirectory()
            root.withdraw()
            if root.filename != '':
                return f'{root.filename}/{day}-{month}-{year}.xls'

    @staticmethod
    def print_title():
        """Prints a nice heading when the program starts."""
        with open('title.txt', 'r') as f:
            for line in f:
                print(line.rstrip())

    def create_sheet_headings(self):
        """Creates the heading rows for the Excel sheet."""
        xlwt.add_palette_colour("colour1", 0x21)
        xlwt.add_palette_colour("colour2", 0x22)
        xlwt.add_palette_colour("colour3", 0x23)
        self.book.set_colour_RGB(0x21, 150, 150, 150)
        self.book.set_colour_RGB(0x22, 247, 150, 70)
        self.book.set_colour_RGB(0x23, 146, 205, 220)
        stats_style = xlwt.easyxf('font: name Arial, bold on, height 160;'
                                  'pattern: pattern solid, fore_colour colour1;'
                                  'borders: top_color black, bottom_color black, right_color black,'
                                  'left_color black, left thin, right thin, top thin, bottom thin')
        style2 = xlwt.easyxf('font: name Arial, bold on, height 160;'
                             'pattern: pattern solid, fore_colour colour2;'
                             'borders: top_color black, bottom_color black, right_color black,'
                             'left_color black, left thin, right thin, top thin, bottom thin;'
                             'align: horiz center;')
        style3 = xlwt.easyxf('font: name Arial, bold on, height 160;'
                             'pattern: pattern solid, fore_colour colour3;'
                             'borders: top_color black, bottom_color black, right_color black,'
                             'left_color black, left thin, right thin, top thin, bottom thin;'
                             'align: horiz center;')
        stats_row = [
            "Date",
            "Heure",
            "Reunion",
            "Hippo",
            "Discip",
            "Course",
            "Partpants"
        ]
        rapports_row = [
            "GAGNANT",
            "GAGNANT PLACE",
            "PLACE",
            "PLACE"
        ]
        couples_row = [
            "CG",
            "CO",
            "CP1",
            "CP2",
            "CP3"
        ]
        """
        COL:   A  B  C  D  E  F  G  H  I  J  K  L  M  N  O  P  Q  R  S  T  U  V  W
        INDEX: 0  1  2  3  4  5  6  7  8  9 10 11 12 13 14 15 16 17 18 19 20 21 22
        For merging cells: write_merge(top_row, bottom_row, left_column, right_column, text, style)
        """

        # A -> G
        # 0 -> 6
        for col_num, col in enumerate(stats_row):
            self.sh.write(0, col_num, col, stats_style)
        for col_num in range(len(stats_row)):
            self.sh.write(1, col_num, "", stats_style)

        # H -> K
        # 7 -> 10
        self.sh.write_merge(
            0, 0, 7, 10,
            "ARRIVEES", style2
        )
        self.sh.write_merge(
            1, 1, 7, 10,
            "LES 4 PREMIERS CHEVAUX ARRIVES", style2
        )

        # L -> O
        # 11 -> 14
        self.sh.write_merge(
            0, 0, 11, 14,
            "RAPPORTS JEUX SIMPLES G P Pour 1€", style3
        )
        for col_num, col in enumerate(rapports_row):
            self.sh.write(1, 11 + col_num, col, style3)

        # P -> T
        # 15 -> 19
        self.sh.write_merge(
            0, 0, 15, 19,
            "COUPLES pour 1€", style3
        )
        for col_num, col in enumerate(couples_row):
            self.sh.write(1, 15 + col_num, col, style3)

        # U
        # 20
        self.sh.write(0, 20, "TRIOS", style3)
        self.sh.write(1, 20, "Désordre", style3)

        # V
        # 21
        self.sh.write(0, 21, "", style3)
        self.sh.write(1, 21, "Ordre", style3)

        # W
        # 22
        self.sh.write(0, 22, "SUPER4", style3)
        self.sh.write(1, 22, "", style3)

        self.sh.write(0, 23, "Ecurie", style3)
        self.sh.write(1, 23, "", style3)

    def save_book(self, filename):
        """Saves the XLS book and checks to see if there are permission errors."""
        while True:
            try:
                self.book.save(filename)
            except PermissionError:
                print(f"\n{filename} could not be saved! You might have it open!")
            else:
                break
            time.sleep(5)  # Waits 5 secs so the loop doesn't spam the console!

    def start_UI(self):
        self.date.get_date()
        self.sh = self.book.add_sheet(
            f'{self.date.year}-{self.date.month}-{self.date.day}',
            # cell_overwrite_ok=True
        )
        self.create_sheet_headings()

    def match_arrays(self, rows1, rows2):
        """Matches two rows based of reunion and course."""
        matched_array = []
        for row1 in rows1:
            for row2 in rows2:
                if row1[0] == row2[0]:
                    matched_array.append(row1[1:] + row2[1:])
        return matched_array

    def order_rows(self, rows):
        ordered_rows = sorted(rows, key=lambda rows: rows[5])
        return sorted(ordered_rows, key=lambda ordered_rows: ordered_rows[2])

    def start_requests(self):
        """Tells Scrapy what requests to deal with first, method automatically called!"""
        self.print_title()
        self.start_UI()
        yield scrapy.Request(
            url=f'{self.BASE_URL}{self.date.year}-{self.date.month}-{self.date.day}',
            callback=self.parse_races
        )

    def closed(self, reason):
        """Tells Scrapy what to do after the spider has closed, method automatically called!"""

        combined_rows = self.match_arrays(self.stats, self.rapport_data)
        combined_rows = self.order_rows(combined_rows)

        # Write stats data to Excel sheet
        for row_num, row in enumerate(combined_rows):
            for col_num, col in enumerate(row):
                # + 2 needed as the first row we want to write on second row first
                self.sh.write(row_num + 2, col_num, col, self.data_style)

        self.file_loc = self.get_file_loc(self.date.day, self.date.month, self.date.year)
        self.save_book(self.file_loc)
        os.startfile(self.file_loc)  # Opens the saved Excel file

    def parse_races(self, response):
        races = response.xpath('//div[@class="yui-g courseLiens  alternate" or '
                               '@class="yui-g courseLiens "]'
                               '//a[normalize-space() = "partants/stats/prono"]/@href').extract()
        # Finds all the urls of the buttons with 'partants/stats/prono' as their text

        rapports = response.xpath('//div[@class="yui-g courseLiens  alternate" or '
                                  '@class="yui-g courseLiens "]'
                                  '//a[normalize-space() = "rapports"]/@href').extract()
        # Finds all the urls of the buttons with 'rapports' as their text

        for race in races:
            # Requsts from all the urls found
            yield scrapy.Request(
                url=f'https://www.geny.com{race}',
                callback=self.parse_pronostics
            )

        for rapport in rapports:
            yield scrapy.Request(
                url=f'https://www.geny.com{rapport}',
                callback=self.parse_rapports
            )

    def parse_pronostics(self, response):
        try:
            """Splits URL up so that date from URL can be easily extracted.
               Example: http://www.geny.com/partants-pmu/
                        2018-07-30-clairefontaine-deauville-pmu-prix-de-la-cote-de-nacre_c991181
            """
            end_url = response.url.split('/')[4].split('-')
            date = f'{end_url[2]}/{end_url[1]}/{end_url[0]}'
        except IndexError:
            # Needed incase the date cannot be found
            date = ''
        hour = response.xpath('//span[@class="infoCourse"]/strong/text()').extract_first()
        if 'h' not in hour:
            # Cheaks if time is in correct format
            hour = ''
        hippo = response.xpath('//div[@id="navigation"]/a[3]/text()').extract_first()
        try:
            reunion = response.xpath(
                '//div[@id="navigation"]'
                '/a[3]/@href'
            ).extract_first().split('#')[1].replace('reunion', 'R')
        except IndexError:
            reunion = ''
        for i in range(1, 6):
            # The main text of info can change in terms of its XPATH reference
            meta_text = response.xpath(f'//span[@class="infoCourse"]/text()[{i}]').extract_first()
            if meta_text.count('-') >= 4 and '\xa0' in meta_text:
                # Reliable way of finding the correct index for the main bit of text
                break
            if i == 6:
                # If by this point the main text hasn't been found, there is an error
                meta_text = ''
        try:
            meta_text = meta_text.split('\xa0')[0]
            meta_text = ''.join([c.strip() for c in meta_text])
            meta_text = meta_text[:-1]
            if meta_text == 'Attelé' or meta_text == 'Monté':
                discipline = 'T'
            elif meta_text == 'Plat':
                discipline = 'P'
            elif meta_text == 'Haies' or meta_text == 'Steeple-chase':
                discipline = 'O'
            else:
                discipline = ''
        except IndexError:
            discipline = ''
        try:
            course = response.xpath('//div[@class="yui-u first nomCourse"]'
                                    '//strong/text()[1]').extract_first().strip()
            # Removed any non digits in the case that something is not scraped correctly
            course = ''.join([c for c in course if c.isdigit()])
        except TypeError:
            course = ''
        partpants = response.xpath('//div[@class="yui-content"]'
                                   '//tbody/tr[last()]/td[1]/text()').extract_first()

        # Stores data for later saving into Excel sheet
        self.stats.append([
            end_url[-1],
            date,
            hour,
            reunion,
            hippo,
            discipline,
            int(course),
            int(partpants)
        ])

    def parse_rapports(self, response):
        end_url = response.url.split('/')[4].split('-')

        try:
            partpants = response.xpath('//table[@id="arrivees"]//tr/td[2]/text()').extract()
            # Cleans text and gets rid of any non digit entries
            partpants = [
                int(partpant.strip())
                for partpant in partpants
                if partpant.strip().isdigit()
            ]
            partpants = max(partpants)
        except IndexError:
            partpants = 0  # Needs to be ten to satisfy the data type of the if statement later on
        try:
            arrivees = response.xpath(
                '//table[@id="arrivees"]//tr[td[1]//text() <= 4]/td[2]/text()'
            ).extract()
            arrivees = [int(arrivee.strip()) for arrivee in arrivees]
            if len(arrivees) > 4:
                arrivees = arrivees[0:4]  # Deals with any draws that may occur
            elif len(arrivees) < 4:
                while len(arrivees) < 4:
                    arrivees.append('')
            # Fills out the arrivees list, if it's not 4 in length
        except (IndexError, ValueError):
            arrivees = ['', '', '', '']

        GAGNANT = ''
        GAGNANT_PLACE = ''
        PLACE1 = ''
        PLACE2 = ''
        CG = ''
        CO = ''
        CP1 = ''
        CP2 = ''
        CP3 = ''
        Désordre = ''
        Ordre = ''
        SUPER4 = ''
        ecurie = ''

        """
        Example First Table:
        +-------+---------+------+
        | Index |         |      |
        +-------+---------+------+
        | 1     | Gagnant | 5,20 |
        +-------+---------+------+
        | 2     | Placé   | 1,80 |
        +-------+---------+------+
        | 3     | Placé   | 3,50 |
        +-------+---------+------+
        | 4     | Placé   | 2,20 |
        +-------+---------+------+
        """

        # Index 1
        GAGNANT = response.xpath(
            '//table[@id="lesSolos"]//tr/td[*//i[text() = "PMU"]]//'
            'tr[*//div[normalize-space() = "Gagnant"]]/td[2]//b/text()'
        ).extract_first()
        # Index 2
        GAGNANT_PLACE = response.xpath(
            '//table[@id="lesSolos"]//tr/td[*//i[text() = "PMU"]]//'
            'tr[*//div[normalize-space() = "Placé"]][1]/td[2]//text()'
        ).extract_first()
        # Index 3
        PLACE1 = response.xpath(
            '//table[@id="lesSolos"]//tr/td[*//i[text() = "PMU"]]//'
            'tr[*//div[normalize-space() = "Placé"]][2]/td[2]//text()'
        ).extract_first()
        # Index 4
        PLACE2 = response.xpath(
            '//table[@id="lesSolos"]//tr/td[*//i[text() = "PMU"]]//'
            'tr[*//div[normalize-space() = "Placé"]][3]/td[2]//text()'
        ).extract_first()

        ecurie = response.xpath(
            '//table[@id="lesSolos"]//tr/'
            'td[*//i[text() = "PMU"]]/div/span/text()'
        ).extract_first()

        if ecurie is not None:
            ecurie = re.sub('\xa0', '', ecurie)  # Cleans encoding
            ecurie = re.sub('Ecurie : ', '', ecurie)  # Makes the "ecuire" numbers easier to read

        """
        Example Second Table:
        +-------+---------+-------+
        | Index |         |       |
        +-------+---------+-------+
        | 1     | Gagnant | 5,20  |
        +-------+---------+-------+
        | 2     | Placé   | 1,80  |
        +-------+---------+-------+
        | 3     | Placé   | 3,50  |
        +-------+---------+-------+
        | 4     | Placé   | 2,20  |
        +-------+---------+-------+
        |       |         |       |
        +-------+---------+-------+
        | 5     | Ordre   | 47,80 |
        +-------+---------+-------+
        """
        if partpants > 7:
            # Index 1
            CG = response.xpath(
                '//table[@id="lesDuos"]//tr/td[*//i[text() = "PMU"]]//'
                'tr[*//div[normalize-space() = "Gagnant"]]/td[2]//b/text()'
            ).extract_first()
            # Index 2
            CP1 = response.xpath(
                '//table[@id="lesDuos"]//tr/td[*//i[text() = "PMU"]]//'
                'tr[*//div[normalize-space() = "Placé"]][1]/td[2]//text()'
            ).extract_first()
            # Index 3
            CP2 = response.xpath(
                '//table[@id="lesDuos"]//tr/td[*//i[text() = "PMU"]]//'
                'tr[*//div[normalize-space() = "Placé"]][2]/td[2]//text()'
            ).extract_first()
            # Index 4
            CP3 = response.xpath(
                '//table[@id="lesDuos"]//tr/td[*//i[text() = "PMU"]]//'
                'tr[*//div[normalize-space() = "Placé"]][3]/td[2]//text()'
            ).extract_first()

        Désordre = response.xpath(
            '//table[@id="lesTrios"]//tr/td[*//i[text() = "PMU"]]/table[1]//tr[2]/td[2]/b/text()'
        ).extract_first()
        # Index 5
        CO = response.xpath(
            '//table[@id="lesDuos"]//tr/td[*//i[text() = "PMU"]]//'
            'tr[*//div[normalize-space() = "Ordre"]]/td[2]//text()'
        ).extract_first()

        if partpants < 8:
            Ordre = response.xpath(
                '//table[@id="lesTrios"]//tr/td[*//i[text() = "PMU"]]//'
                'tr[*//div[normalize-space() = "Ordre"]]/td[2]//text()'
            ).extract_first()
            SUPER4 = response.xpath(
                '//table[@id="lesQuartos"]//tr/td[*//i[text() = "PMU"]]//'
                'tr[*[normalize-space() = "Super 4"]]/td[2]//text()'
            ).extract_first()

        data = [
            GAGNANT,
            GAGNANT_PLACE,
            PLACE1,
            PLACE2,
            CG,
            CO,
            CP1,
            CP2,
            CP3,
            Désordre,
            Ordre,
            SUPER4,
            ecurie
        ]

        data = [
            ''
            if d is None
            else d
            for d in data
        ]
        # Converts XPaths that couldn't be found to empty strings
        data = [d.strip() for d in data]
        # Cleans all strings

        data = [re.sub(' €', '', d) for d in data]
        # Gets rid of Euro sign and spaces
        data = [re.sub(',00', '', d) for d in data]
        # Replaces for example 3,00 with 3
        data = [int(d) if d.isdigit() else d for d in data]
        # Converts strings that are ints, to ints, so Excel doesn't show errors

        # Stores data for later saving into Excel sheet
        self.rapport_data.append([end_url[-1]] + arrivees + data)
