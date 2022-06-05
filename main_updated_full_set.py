import module1
import module2
import pdfreader
import os
import glob
from pathlib import Path
import xlwt
from xlwt import Workbook

topics_to_keywords = { 
    'economy': [ "οικονομια", "επιδοτηση", "επιτοκιο", "φορο", "χρεος", "εε", "φπα" ],
    'fiscal tax policy': [ "ελλειμμα", "προυπολογισμος", "επιδοματ", "φόρος", "φόρος προστιθέμενης αξίας", "φπα", "ειδικός φόρος κατανάλωσης", "δημοσια έσοδα", "έσοδα ιδιωτικοποιήσεων", "ιδιωτικοποιηση", "εφορια", "ανεξάρτητη αρχή δημοσιων εσοδων", "ααδε" ],
    'fiscal debt policy': [ "δημοσιες δαπαν", "ελλειμμα", "προυπολογισμος", "επιδοματ", "κρατικες δαπανες", "πρωτογενεις δαπανες", "αμυντικες δαπανες", "δημοσιες επενδυσεις", "προυπολογισμος", "εθνικο χρεος", "δημοσιο χρεος", "πληρωμες μεταβιβασης", "δημοσια καταναλωση", "δημοσιο οφελος", "επιδομα", "χρεοκοπια", "χρεοκοπια χωρας" ],
    'monetary policy': [ "δημοσιες δαπαν", "επιτοκιο", "κοστος χρηματος", "νομισματικη πολιτικη", "ποσοτικη χαλαρωση" ],
    'currency policy': [ "συναλλαγματική ισοτιμία", "δραχμη", "ευρωζωνη", "ανατιμηση νομισματος", "ανατιμησ", "υποτιμηση νομισματος", "υποτιμησ", "εθνικο νομισμα", "οικονομικη και νομισματική ένωση", "ονε", "Grexit", "εξοδος απο ευρωζ" ],
    'banking policy': [ "τραπεζα", "τραπεζικο συστημα", "τραπεζικος τομεας", "δανειο", "καταθεσεις", "διατραπεζικη αγορα", "επιτοκιο δανεισμου", "euribor", "επιτοκιο καταθεσεων" ],
    'pension policy': [ "συνταξη", "εφαπαξ συνταξη", "συστημα συνταξιοδοτικής ασφάλισης", "ασφαλιστικό ταμείο", "ιδρυμα Κοινωνικών Ασφαλίσεων", "ικα", "κοινωνική ασφάλιση", "ρήτρα μηδενικού ελλείμματος", "συνταξιοδοτική μεταρρύθμιση", "ασφαλιστική εισφορά", "κεφαλαιοποιητικό συνταξιοδοτικό σύστημα", "συνταξιοδοτικό σύστημα αναδιανεμητικο", "επιδομα ανεργιας", ],
    'government policy': [ "νομοθεσια", "νομοσχεδιο", "τραπεζα της ελλαδος", "ττε", "κεντρικη τραπεζα", "μεταρρυθμιση", "διαρθρωτικες αλλαγες", "πρωθυπουργος", "μεγαρο μαξιμου", "ελλειμμα", "κανονιστικο πλαισιο", "επιτροπη κεφαλαιαγορας", "επιτροπη ανταγωνισμου", "συμβουλιο της επικρατειας", "στε" ],
    'defence policy': [ "στρατος", "στρατιωτικες δαπανες", "αμυντικες δαπανες", "εξοπλιστικο προγραμμα", "στρατιωτικ θητεια", "οπλισμο", "εθνικη αμυνα", "πολεμικη αεροπορια", "πενταγωνο", "πολεμικο ναυτικο", "στρατος ξηρας", "πεζικο", "εαβ" ],
}

topics_stats = {}

for topic in topics_to_keywords:
    keywords = topics_to_keywords[topic]
    topics_stats[topic] = {}
    for keyword in keywords:
        topics_stats[topic][keyword] = 0

def get_frequency_by_keyword(text, keywords):
    stats = {}
    text = text.lower()

    for keyword in keywords:
        stats[keyword] = 0

    for keyword in keywords:
        stats[keyword] += text.count(keyword)

    return stats;

def get_frequency_by_topic_and_keyword(text, topics_to_keywords):
    stats = {}
    text = text.lower()

    find_to_replace = {
        'ά': 'α',
        'έ': 'ε',
        'ή': 'η',
        'ί': 'ι',
        'ό': 'ο',
        'ύ': 'υ',
        'ώ': 'ω',
    }

    for f, r in find_to_replace.items():
        text = text.replace(f, r)

    for topic in topics_to_keywords:
        frequency_by_keyword = get_frequency_by_keyword(text, topics_to_keywords[topic])
        stats[topic] = {
            'frequency_by_keyword': frequency_by_keyword,
            'total': 0
        }

        for keyword in frequency_by_keyword:
            stats[topic]['total'] += frequency_by_keyword[keyword]

    return stats


def write_excel(stats):
    wb = Workbook()
  
    # add_sheet is used to create sheet.
    sheet1 = wb.add_sheet('Sheet 1')

    col = 1

    for starting_year in stats:
        frequency_by_topic_and_keyword = stats[starting_year]
        sheet1.write(0, col, starting_year)
        col += 1

    row = 1

    for starting_year in stats:
        frequency_by_topic_and_keyword = stats[starting_year]
        for topic in frequency_by_topic_and_keyword:
            frequency_by_keyword = frequency_by_topic_and_keyword[topic]['frequency_by_keyword']
            for keyword in frequency_by_keyword:
                sheet1.write(row, 0, keyword)
                row += 1
        break

    col = 1
    row = 1

    for starting_year in stats: 
        frequency_by_topic_and_keyword = stats[starting_year]

        for topic in frequency_by_topic_and_keyword:
            frequency_by_keyword = frequency_by_topic_and_keyword[topic]['frequency_by_keyword']
            for keyword in frequency_by_keyword:
                frequency = frequency_by_keyword[keyword]
                sheet1.write(row, col, frequency)
                row += 1

        row = 1
        col += 1
    
    wb.save('statistika.xls')


directory_to_search = r"C:\Users\Christ05\.VS_code_PhD_data\Πρακτικά Βουλής_2000_2020_in pdf"

pathname = directory_to_search + "/**/*.pdf"
pdf_file_paths = glob.glob(pathname, recursive=True)

stats = {}

for pdf_file_path in pdf_file_paths:
    pdf_file_info = pdfreader.read_pdf_info(pdf_file_path)
    pdf_text = pdf_file_info["text"].lower()
    filename_parts = Path(pdf_file_path).stem.split('_')
    starting_year = filename_parts[0]

    frequency_by_topic_and_keyword = get_frequency_by_topic_and_keyword(pdf_text, topics_to_keywords)
    stats[starting_year] = frequency_by_topic_and_keyword

    write_excel(stats)


print(stats)
