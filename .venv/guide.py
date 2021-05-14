from fuzzywuzzy import process
from xlrd import open_workbook
import xlrd
from xlsxwriter.utility import xl_rowcol_to_cell
import numpy as np
from termcolor import colored
import pyfiglet
from pyfiglet import Figlet
print("")
ascii_banner = pyfiglet.figlet_format("Gemstone Cutter's Guide")
print(ascii_banner)
print("")


str2Match = input("Which stone's features do you want to find?: ")

strOptions = ["Hematite","Chloastrolite","Feldspar","Goldstone","Psilomelane","Petosky stone","Unakite","Wonderstone","Cinnabar","Proustite","Pyrargyrite","Cuprite","Rutile","Brookite","Anatase","Diamond","Fabulite","Stibiotantalite","Sphalerite","Crocoite","Wulfenite","Tantalite","Linobate","Manganotantalite","Cubic zirconia (CZ)","Mimetite","Phosgenite","Senarmontite","Boleite","Zincite","Cassiterite","Simpsonite","Gadolinium gallium garnet (GGG)","Sulfur","Bayldonite","Scheelite","Andradite garnet","Anglesite","Uvarovite garnet","Purpurite","Sphene (titanite)","Yttrium aluminum garnet (YAG)","Zircon","Cerussite","Gahnite","Spessartite garnet","Painite","Monazite","Almandine garnet","Gadolinite","Ruby (corundum)","Sapphire (corundum)","Benitoite","Shattuckite","Chrysoberyl","Periclase","Scorodite","Staurolite","Grossular garnet","Chambersite","Hessonite garnet","Epidote","Pyroxmangite","Azurite","Pyrope garnet","Hodgkinsonite","Taaffeite","Rhodonite","Gahnospinel","Spinel","Kyanite","Adamite","Diaspore","Serendibite","Sapphirine","Aegirine-augite","Idocrase (vesuvianite)","Tanzanite","Neptunite","Willemite","Rhodizite","Triphylite","Lithiophilite","Dumortierite","Legrandite","Hypersthene","Parisite","Clinozoisite","Sinhalite","Lawsonite","Diopside","Bustamite","Kornerupine","Hiddenite","Kunzite","Boracite","Axinite","Malachite","Sillimanite","Jadeite","Peridot","Ludlamite","Enstatite","Euclase","Phenakite","Dioptase","Jet","Eosphorite","Spurrite","Jeremejevite","Barite","Siderite","Danburite","Clinohumite","Apatite","Andalusite","Friedelite","Smithsonite","Datolite","Celestite","Tourmaline","Actinolite","Hemimorphite","Lazulite","Prehnite","Gaspéite","Turquoise","Topaz","Sugilite","Sogdianite","Brazilianite","Rhodochrosite","Odontolite","Nephrite","Pectolite (larimar)","Montebrasite","Phosphophyllite","Meliphanite","Eudialyte","Chondrodite","Catapleiite","Wardite","Herderite","Colemanite","Howlite","Zektzerite","Amblygonite","Ekanite","Anhydrite","Augelite","Emerald (beryl)","Aquamarine (beryl)","Variscite","Beryl (precious)","Tremolite","Vivianite","Serpentine","Larbradorite","Hambergite","Pyrophyllite","Muscovite","Beryllonite","Charoite","Amethyst (quartz)","Aventurine (quartz)","Rose (quartz)","Citrine (quartz)","Prasiolite (quartz)","Smoky (quartz)","Rock crystal (quartz)","Andesine","Cordierite","Oligoclase","Talc","Scapolite","Petrified Wood","Jasper","Amber","Ivory","Apophyllite","Tiger’s eye","Aragonite","Agate","Chalcedony","Chrysoprase","Moss agate","Sepiolite","Witherite","Milarite","Nepheline","Sunstone","Amazonite","Pearl","Ammolite","Strontianite","Gypsum","Orthoclase","Sanidine","Moonstone","Pollucite","Carletonite","Stichtite","Thomsonite","Magnesite","Scolecite","Leucite","Mesolite","Dolomite","Petalite","Lapis lazuli","Haüyne","Tugtupite","Cancrinite","Celluloid","Ulexite","Yugawaralite","Whewellite","Kurnakovite","Inderite","Calcite","Coral","Moldavite","Natrolite","Sodalite","Analcime","Thaumasite","Creedite","Chrysocolla","Obsidian","Gaylussite","Glass","Fluorite","Sellaite","Opal"]
Ratios = process.extract(str2Match,strOptions, limit=10)
print("")
print("")
print("POSSIBLE MATCHES")
print("")
print(Ratios)

print("")
print("")
print("")
print('If first option above is correct, JUST PRESS ENTER')

str2Match_2 = input("If it is not, then type the name again by looking up above list: ")

if str2Match_2=="":
    Ratios = process.extractOne(str2Match, strOptions)

else:
    str2Match = str2Match_2
    Ratios = process.extractOne(str2Match, strOptions)

arr = np.asarray(Ratios)
fla_arr = arr.flatten()
str2Match = Ratios[0]

#For Example: workbook1 = xlrd.open_workbook(r"C:\Users\user\Desktop\gemstone-cutter-guide\.venv\Gem-List.xls")
workbook1 = xlrd.open_workbook(r"Path of the Gem File in the folder.")
sheet1 = workbook1.sheet_by_index(0)


#For Example: for sh in xlrd.open_workbook(r"C:\Users\user\Desktop\gemstone-cutter-guide\.venv\Gem-List.xls").sheets():
for sh in xlrd.open_workbook(r"Same thing again").sheets():
    for row in range(sh.nrows):
        for col in range(sh.ncols):
            myCell = sh.cell(row, col)
            if myCell.value == str2Match:
                print('___________________________________________________________________________________________')
                #print('Found!')
                #print(xl_rowcol_to_cell(row, col))
               ## print ("Gemstone	Refractive Index	Double Refraction")
                custom_fig = Figlet(font='Doom')
                print(custom_fig.renderText(sheet1.row_values(row)[0]))
                print('___________________________________________________________________________________________')
                print('FACETING')
                print("Gemstone:          " + str(sheet1.row_values(row)[0]))
                print("Refractive Index:  " + str(sheet1.row_values(row)[1]))
                print("Double Refraction: " + str(sheet1.row_values(row)[2]))
                print('___________________________________________________________________________________________')
                print('POLISHING')
                print("Canvas:   " + str(sheet1.row_values(row)[3]))
                print("Phenolic: " + str(sheet1.row_values(row)[4]))
                print("Felt:     " + str(sheet1.row_values(row)[5]))
                print("Leather:  " + str(sheet1.row_values(row)[6]))
                print("Muslin:   " + str(sheet1.row_values(row)[7]))
                print('___________________________________________________________________________________________')
                quit()


