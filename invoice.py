#!/usr/bin/python
import locale
import os
import subprocess
import sys
from appy.pod.renderer import Renderer
from datetime import date, timedelta, datetime
from dateutil.relativedelta import relativedelta
from decimal import Decimal
from babel.dates import format_date
import pytils.numeral
from num2words import num2words



DATE_FORMAT = '%d.%m.%Y'


# Allow unicode in file names.
reload(sys)
sys.setdefaultencoding('utf8')

today = date.today()
invoice_date = today - timedelta(days=today.day)


def date_to_str(d):
    return d.strftime(DATE_FORMAT)

def str_to_date(d):
    return datetime.strptime(d, DATE_FORMAT).date()

default_locale = locale.getlocale()

def date_to_month_year(date, loc):
    locale.setlocale(locale.LC_TIME, loc)
    s = date.strftime('%B %Y')
    locale.setlocale(locale.LC_TIME, default_locale)
    return s

def date_to_text(date, loc):
    return format_date(date, 'dd MMMM YYYY', loc)
  

INPUTS = (
    # ('name', 'default value')
    ('contract_start', date_to_str(
        today - relativedelta(months=1) - timedelta(today.day - 1))),
    ('contract_end', date_to_str(
        today - timedelta(today.day))),
    ('invoice_date', date_to_str(today)),
    ('items', (
        ('number', lambda number, inputs: str(number)),
        ('name', 'Item %i'),
        ('amount', '0.00'),
    )),
)

CALCS = (
    # ('name', function)
    ('annex_number', lambda data:
        ''.join(reversed(data['contract_start'].split('.')))),
    ('protocol_number', lambda data:
        ''.join(reversed(data['invoice_date'].split('.')))),
    ('invoice_number', lambda data:
        ''.join(reversed(data['invoice_date'].split('.')))),
    ('total_amount', lambda data:
        format(sum(Decimal(item['amount']) for item in data['items']), '.2f')),
    ('duration_before', lambda data:
        date_to_str(datetime.strptime(data['invoice_date'], DATE_FORMAT) +
                    relativedelta(months=1))),
)

DOCS = (
    # template file, output file
    ('templates/contract-annex-template.odt', 'contract-annex-%(annex_number)s.odt'),
    ('templates/protocol-template.odt', 'protocol-%(protocol_number)s.odt'),
    ('templates/invoice-template.odt', 'invoice-%(invoice_number)s.odt'),
)


def read_inputs(input_list, current_number=0):
    inputs = {}
    for key, value in input_list:
        if isinstance(value, tuple):
            query = 'Number of %s (0) : ' % key.replace('_', ' ')
            number = int(raw_input(query) or 0)
            inputs[key] = []
            for i in range(number):
                inputs[key].append(read_inputs(value, i + 1))
        elif callable(value):
            inputs[key] = value(current_number, inputs)
        else:
            value = value.replace('%i', str(current_number))
            query = '%s (%s): ' % (key.replace('_', ' ').capitalize(), value)
            inputs[key] = raw_input(query).strip() or value

    return inputs


def shellquote(s):
    return "'" + s.replace("'", "'\\''") + "'"


data = read_inputs(INPUTS)
calculateds = {}
for key, func in CALCS:
    data[key] = func(data)

print 'Gethered data:'
print data

print 'Rendering the documents...'
for template_file, output_file in DOCS:
    output_file = output_file % data
    print '%s -> %s' % (template_file, output_file)
    template_file = os.path.abspath(template_file)
    output_file = os.path.abspath(output_file)
    if os.path.isfile(output_file):
	s_question = 'File %s already exists. Do you want to overwrite it? (y): '
	if (raw_input(s_question % output_file) or 'y').lower() != 'y':
	    exit()
	os.remove(output_file)

    dirname = os.path.dirname(output_file)
    if not os.path.isdir(dirname):
	os.mkdir(dirname)
    
    print "    rendering to ODT"
    renderer = Renderer(template_file, data, output_file)
    renderer.run()

    print "    converting to PDF"
    subprocess.check_call(
        "HOME=%s /usr/bin/soffice --headless --invisible --convert-to pdf "
        "%s --outdir %s > /dev/null 2>&1" %
        (shellquote(os.path.abspath('.')), shellquote(output_file), os.path.abspath(dirname)),
        shell=True)
    print "    converting to JPG"
    subprocess.check_call(
        "HOME=%s /usr/bin/convert -density 300 "
        "%s %s > /dev/null 2>&1" %
        (shellquote(os.path.abspath('.')),
	 shellquote(os.path.splitext(output_file)[0] + '.pdf'),
	 shellquote(os.path.splitext(output_file)[0] + '.jpg')),
        shell=True)
print 'Done!'
