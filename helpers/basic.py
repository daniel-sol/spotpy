"""Author dbs: helper functions for use in spotfire scripts"""
import logging
import sys
import re


def basic_path(path_list):

    """Returns a basic path, to be used for file imports
    """
    path = r'\\'.join(path_list)


    return path


def init_logging(name, loglevel=False, file_name=None, full=False):

    """
     Init of logger instance
     args:
        name: string, name that will appear when logger is activated

        loglevel: boolean, defaults to true, if otherwise

        file_name: string or None (default), name of file handle for logger


      returns logger: python logging instance

      raises: ValueError: if loglevel is not among the predefined
              logging level types

    """

    if full:

        format = ('%(asctime)s: %(name)s - %(funcname)s %(levelname)s ' +
                  '- %(message)s')

    else:

        format = '%(asctime)s: %(message)s'

    dateformat = '%m/%d/%Y %I:%M:%S'
    if not loglevel:

        logger = logging.getLogger(name)
        logger.addHandler(logging.NullHandler())

    else:

        # assuming loglevel is bound to the string value obtained from the
        # command line argument. Convert to upper case to allow the user to
        #use both lower and upper case
        numeric_level = getattr(logging, loglevel.upper(), None)

        if not isinstance(numeric_level, int):

            raise ValueError('Invalid log level: %s' % loglevel)

        logging.basicConfig(level=numeric_level, format=format, datefmt=dateformat)
        logger = logging.getLogger(__name__)

        #Adds file handle to logger if file_name is specified
        if file_name is not None:

            try:
                fhandle = logging.FileHandler(file_name)
                formatter = logging.Formatter(format)
                fhandle.setFormatter(formatter)
                logger.addHandler(fhandle)

            except IOError as ioerr:

                raise ioerr

    return logger


def item_dict():
    """returns dictionary with items to find in spotfire
       returns items: dictionary of the data types
    """

    items = {'page': 'pages', 'table': 'tables',
             'viz': 'vizualisation', 'column': 'columns'}
    return items


def join_list(jlist, joiner=', '):

    """Joins a list into a list

    args:
      jlist (list of strings): list to be joined
      joiner (str): a string to be used as a joint
    returns join_string: str

    """
    if len(jlist) == 0:
        jlist = '[]'
    else:
        jlist = joiner.join(jlist)
    return jlist


def controlled_list(input_list):

    """converts something that might be something else to list
    args:

    input_list (something):

    returns output_list
    """
    output_list = input_list

    if not isinstance(input_list, list):

        dummy_list = []
        dummy_list.append(input_list)
        output_list = dummy_list
        print('Converting')
    print('Before return')
    print(output_list)
    return output_list


def create_line(line_list):
    """Makes a string from a list

    args:
       line_list(list): list to make line from

    returns: line (str)

    """
    line = join_list(line_list, '\t') + '\r\n'

    return line


def list_string(join_list):

    """Returns a list as a string like so [el1, el2 ...]
    args:
    join_list (list)
    returns joined_list (str)
    """
    joined_list = '[{}]'.format(join_list, join_list)
    return joined_list


def get_list_entry(inlist, i):

    """Fetches entry from list by index, returns empty string if i not in list

    args:

      inlist (list):
      i (int): entry to get from list
    returns out (list entry or ' ')
    """
    logger = init_logging(__name__ + '_get_list_entry')
    out= ' '
    try:

        out = inlist[i]

    except IndexError:

        logger.debug('Nothing to extract at %i', i)

    return out


def set_document_property(doc, prop_name, value_name):

    """Setting document property
    args:
    doc (Spotfire document instance): document to read from
      prop_name (str): name of property
      value_name (str): value
    """
    from Spotfire.Dxp.Data import DataProperty
    from Spotfire.Dxp.Data import DataType
    from Spotfire.Dxp.Data import DataPropertyClass

    logger = init_logging(__name__ + '.set_document_property')
    logger.debug('Trying to set {} to {}'.format(prop_name, value_name))
    try:
        doc.Properties[prop_name] = value_name
    except KeyError:
        logger.info('%s does not exist will create as a string', prop_name)
        attr = DataProperty.DefaultAttributes
        prop = DataProperty.CreateCustomPrototype(prop_name, DataType.String, attr)
        doc.Data.Properties.AddProperty(DataPropertyClass.Document, prop)
        doc.Properties[prop_name] = value_name

    message = '{} set to {}'.format(prop_name, value_name)
    heading = 'Setting document property'
    ok_message(message, heading)


def part_expression(ex_name, stat_type):
    """ makes part of expression to be used on chart axis
        ex_name (str): curve name
        stat_type (str): name of statistical type
    """
    if stat_type =='':
        start = '['
        end = ']'
    else:
        start = '(['
        end = '])'
    expression = stat_type + start + ex_name + end
    return expression


def expression_maker(ex_names, stat_type):
    """ makes expression to be used on chart axis
        args:
        ex_names (list): curve names
        stat_type (str): name of statistical type
    """
    if len(ex_names) == 1:
        expression = part_expression(ex_names[0], stat_type)
    else:
        current_part = ex_names.pop(-1)
        expression = (expression_maker(ex_names, stat_type) + ','
                      + part_expresion(current_part, stat_type))

    return expression


def create_expression(ex_names, stat_types):
    """ makes expression to be used on chart axis
        args:
        ex_names (list): curve names
        stat_types (list): name of statistical type
    """
    expression = ''
    if not isinstance(ex_names, list):
        ex_names = list(ex_names)

    if not isinstance(stat_types, list):
        stat_types = list(stat_types)

    expression = expression_maker(ex_names, stat_types.pop(0))
    for stat_type in stat_types:
        expression = expression + ',' + expression_maker(ex_names, stat_types)
