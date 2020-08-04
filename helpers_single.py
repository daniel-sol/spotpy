"""Author dbs: helper functions for use in spotfire scripts"""
import logging
import sys
import re



def join_list(jlist, joiner=', '):

    """Joins a list into a list

    args:
      jlist (list of strings): list to be joined
      joiner (str): a string to be used as a joint
    returns join_string: str

    """
    jlist = joiner.join(jlist)
    return jlist


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

        format = '%(asctime)s: %(name)s - %(levelname)s - %(message)s'

    else:

        format = '%(asctime)s: %(message)s'

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


def announce_no_data(name, valid_names, data_type):

    """Announcing that no data is found in a message box

       args:

           name     : str
           valid_data: list
           data_type :  str, name of data type not found

       raises: AttributeError if data_type is not in valid_types
    """

    logger = init_logging(__name__)
    valid_types = ['Page', 'Table', 'Chart', 'Column']

    if data_type not in valid_types:

        raise AttributeError('No such data type in spotfire project')

    else:
        heading = 'Data does not exist!'
        message = ('{} with name does not exist,\n please choose among the' +
                   ' following').format(data_type,
                                        name, join_list(valid_names,'\n'))

        ok_message(message, heading)
        logger.warning(message)


def ok_message(message, heading='WARNING'):

    """ Creates information box
        args:

            message: str, what is what is displayed in box

            heading: str, what is displayed as heading 'name' of box

    """
    import clr
    clr.AddReference('System.Windows.Forms')
    from System.Windows import Forms
    Forms.MessageBox.Show(message, heading, Forms.MessageBoxButtons.OK)


def yes_no_message(message, heading):

    """ Creates decision box with yes and no, returns boolean

        args:

            message (str): text displayed in box

            heading (str): text displayed as heading 'name' of box

            returns:
                  answer: boolean

    """
    logger = init_logging(__name__)
    import clr
    clr.AddReference('System.Windows.Forms')
    from System.Windows import Forms
    answer = False
    reply = Forms.MessageBox.Show(message, heading, Forms.MessageBoxButtons.YesNo)
    if reply == Forms.DialogResult.No:

        logger.debug('User pressed no')

    else:
        answer = True
        logger.debug('User pressed yes')

    return answer


def get_page_names(doc):

    """Fetches names of the different pages in a spotfire project

    returns:
        names: list of strings, names of the different pages

    """
    logger = init_logging(__name__)
    names = []
    logger.debug('Looping through pages')
    for page in doc.Pages:

        names.append(page.Title)

    logger.debug('Returning table')
    logger.debug(names)

    return names


def get_table_names(doc):

    """Fetches the data table names in a spotfire project
    args:
       doc (Spotfire document instance): document to find tables from
    returns:
        names: list of strings, names of the different tables

    """
    logger = init_logging(__name__)
    names = []
    logger.debug('Looping through tables')
    for table in doc.Data.Tables:

        names.append(table.Name)

    logger.debug('Returning table')
    return names


def get_visual_names(page, viz_type=None):
    """Fetches the visualization names on a page in a spotfire project

    args:
        page:Spotfire document page instance

    returns:
        visuals: dictionary of strings, names of the different visualizations


    """
    # import Spotfire.Dxp.Application.Visuals as viz
    visuals = {}
    for visual in page.Visuals:
        name = visual.Title
        viz_id = visual.TypeId.Name.replace('Spotfire.', '')
        visuals[str(name)] = viz_id

        if viz_type is not None:

            visuals = {key:value for key, value in visuals\
                       if re.match(r'.*{}'.format(viz_type), value)}
    return visuals


def get_visual(doc, page_name, viz_name, check=False):
    """ Getting visual from page
    args:
      doc (Spotfire document instance): Document to work with
      page_name (str): name of page

    """
    from Spotfire.Dxp.Application.Visuals import VisualContent
    page = get_page(doc, page_name)
    viz_names = get_visual_names(page)
    # print(viz_names)
    viz = None
    if not viz_names:
        message = 'No visuals on page {}'.format(page_name)
        ok_message(message)

    else:

        for vis in page.Visuals:
            if vis.Title == viz_name:
                viz = vis
                viz = viz.As[VisualContent]()
        if viz is None and check:
            message = ('There is no visual on page {} ' +
                       'called {}').format(page_name, viz_name)

            message += ('\nAvailable visualizations are' +
                        '[{}]').format(join_list(viz_names))
            ok_message(message)

    return viz


def get_chart_names(page):

    """Fetches the chart names on a page in a spotfire project

    args:

       page: Spotfire document page instance

       returns:

           chart_names: dictionary with keys and values as strings

    """

    chart_names = get_visual_names(page)

    return chart_names


def get_table(doc, table_name='Active'):

    """Fetches a specific datatable
    args:
      doc (Spotfire document instance): Document to work with
      table_name (str): name of table

    returns:
        table: None or Spotfire data table

    raises: AttributeError: if table with name does not exist

    """
    logger = init_logging(__name__)
    logger.debug('Fetching table with name %s', table_name)
    table = None
    if table_name == 'Active':

        table = doc.ActiveDataTableReference

    else:
        try:
            table = doc.Data.Tables[table_name]
        except IndexError:

            valid_tables = get_table_names(doc)
            announce_no_data(name, valid_tables, 'Table')

    return table


def add_viz_to_page(doc, viz_type, viz_name, table_name, page_name='Active'):
    """Adds vizualisation to page
    args:
    doc (Spotfire document instance): document to read from
    page_name (str): name of page
    viz_type(str): type of visual
    viz_name (str): name of visual
    returns viz (spotfire visual)
    """
    from Spotfire.Dxp.Application.Visuals import BarChart, BoxPlot, ScatterPlot
    from Spotfire.Dxp.Application.Visuals import  LineChart, HtmlTextArea
    from Spotfire.Dxp.Application.Visuals import Visualization

    viz_dict = {'barchart': BarChart, 'boxplot': BoxPlot,
                'scatter': ScatterPlot, 'linechart': LineChart,
                'textarea': HtmlTextArea}

    viz = None
    if viz_type not in viz_dict:

        heading = 'WARNING!'
        message = ('{} is not a valid visual choose among ' +
                   '[{}]').format(viz_type, join_list(sorted(viz_dict)))

        ok_message(message, heading)

    else:

        page = get_page(doc, page_name)
        viz = get_visual(doc, page_name, viz_name)

        if viz is None:
            viz = page.Visuals.AddNew[viz_dict[viz_type]]()
            message = 'Making visual {} on page {}'.format(viz_name,
                                                           page_name)
            heading = "INFO"
            ok_message(message, heading)

        print(viz)
        viz.Data.DataTableReference = get_table(doc, table_name)
        try:
            viz.Title = viz_name
            viz.ShowTitle = True
        except AttributeError:
            print('Cannot set title')

        print('Created {}'.format(viz_name))
    return viz


def make_viz(doc, viz_type, viz_name, table_name, x_name, y_names, page_name='Active'):

    """Adds vizualisation to page
    args:
    doc (Spotfire document instance): document to read from
    page_name (str): name of page
    viz_type(str): type of visual
    viz_name (str): name of visual
    returns viz (spotfire visual)
    """
    viz = add_viz_to_page(doc, viz_type, viz_name, table_name, page_name)
    viz.Title = viz_name
    if x_name == 'DATE':

        x_expression = 'BinByDateTime([DATE],"Year.Month",2)'

    else:
        x_expression = "[" + x_name + "]"
    viz.XAxis.Expression = x_expression

    color_expression = '<[Axis.Default.Names]>'
    if not isinstance(y_names, list):

        y_expression = y_expression + 'Avg (' + y_names + ')'

    else:

        y_expression = ' Avg (' + y_names.pop(0) + ')'

        for y_name in y_names:
            y_expression = y_expression + ', Avg (' + y_name + ')'

    viz.YAxis.Expression = y_expression
    viz.ColorAxis.Expression = color_expression

    viz.Legend.Visible = True


def get_table_column_names(doc, table_name):

    """Fetches the column names from a data table
    args:
    doc (Spotfire document instance): document to read from
    table_name (str): name of table
    returns. col_names (list): names of the columns
    """
    table = get_table(doc, table_name)
    col_names = [c.Name for c in table.Columns]
    return col_names


def check_col_names(doc, table_name , name_list):
    """Checks col_names in list

    """


def get_well_names(doc, table_name):

    """Fetches well_names from columns in a table
    args:
    doc (Spotfire document instance): document to read from
    table_name (str): name of table
    returns. well_names (list): names of wells
    """
    logger = init_logging(__name__ + '_get_well_names')
    logger.debug('Fetching table with name %s', table_name)
    col_names = get_table_column_names(doc, table_name)
    well_names = []
    for col_name in col_names:
        m = re.match(r'W[^:]+.(.*)', col_name)
        if m:
            well_name = m.group(1)
            if well_name not in well_names:
                well_names.append(well_name)
        else:
            logger.debug(col_name + " does not match")

    return well_names


def get_list_entry(inlist, i):

    """ Fetches entry from a list by index

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

def get_html(doc, page_name):

    from Spotfire.Dxp.Application.Visuals import HtmlTextArea
    viz = get_visual(doc, page_name, 'Text Area')


    html = viz.As[HtmlTextArea]().HtmlContent
    return html


def make_vector_and_well_tables(doc, in_table_name):


    """Makes tables with names of vectors and wells
    args
    doc (Spotfire document instance): document to read from
    in_table_name (str): name of table to extract from

    """
    vector_names = get_vector_names(doc, in_table_name)
    well_names = get_well_names(doc, in_table_name)

    vector_table_name = 'Vectors'
    make_table(doc, vector_table_name, vector_table_name , vector_names)
    well_table_name = 'Wells'
    make_table(doc, well_table_name, well_table_name , well_names)


def make_vector_and_well_table(doc, in_table_name,
                               out_table_name='Vectors_and_wells'):

    """Makes table with names of vectors and wells
    args
    doc (Spotfire document instance): document to read from
    in_table_name (str): name of table to extract from
    out_table_name (str): name of table to make

    """
    vector_names = get_vector_names(doc, in_table_name)
    well_names = get_well_names(doc, in_table_name)

    length = 0

    for entry in (vector_names, well_names):

        entry_length = len(entry)
        if entry_length > length:

            length = entry_length

    header = ['Vectors', 'Wells']
    text_list = []
    i = 0
    while i <=length:

        text_list.append([get_list_entry(vector_names, i),
                          get_list_entry(well_names, i)])
        i +=1

    print(text_list)

    make_table(doc, out_table_name, header, text_list)


def get_vector_names(doc, table_name):
    """Fetches vector names from columns in a table
    args:
    doc (Spotfire document instance): document to read from
    table_name (str): name of table
    returns. vector_names (list): names of vectors
    """
    col_names = get_table_column_names(doc, table_name)
    vector_names = []
    for col_name in col_names:
        if re.match(r'(date|time)', col_name.lower()):
            continue
        extract_name = col_name.split(':')[0]
        extract_name = extract_name[1:]
        if extract_name.endswith('H'):
            extract_name = extract_name[:-1]

        if extract_name not in vector_names:
            vector_names.append(extract_name)
    return vector_names


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


def set_document_properties(doc, table_name):

    """Setting the document properties wells and vectors
       doc (Spotfire document instance): document to read from
    """
    set_document_property(doc, 'wells',
                          join_list(get_well_names(doc, table_name), ';'))
    #set_document_property(doc, 'well_list', get_well_names(doc, table_name))
    set_document_property(doc, 'vectors',
                          join_list(get_vector_names(doc, table_name), ';'))


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


def table_text_writer(list_input):

    """Writes table input

    args:

    list_input (list of lists): input to be written

    returns: stream
    """
    from System.IO import  StreamWriter, MemoryStream, SeekOrigin


    list_input = controlled_list(list_input)
    print(list_input)
    print('--------')
    header = controlled_list(list_input.pop(0))
    length_of_header = len(header)
    head_line = create_line(header)

    # if only a header line
    if not list_input:
        text = ' ' * (length_of_header)

    else:
        text =''

        print('Loop')
        for i, text_list in enumerate(list_input):

            text_list = controlled_list(text_list)
            print(text_list)
            if len(text_list) != length_of_header:

                message = ('Line {} in text: [{}] does not have ' +
                           ' same length as header [{}], will be ' +
                           'skipped').format(i, join_list(text_list),
                                             join_list(header))
                ok_message(message)
            else:

                text += create_line(text_list)

    text = head_line + text
    stream = MemoryStream()
    writer = StreamWriter(stream)
    writer.Write(text)
    writer.Flush()
    stream.Seek(0, SeekOrigin.Begin)
    return stream


def delete_columns(doc, table_name, column_names):
   """Removes columns from data table
    args:
    doc (Spotfire document instance): document to read from
    table_name (str): name of table
    col_list (list): names of the columns to be deleted
    """
   table = get_table(doc, table_name)


   message = 'Remove columns [{} from table {}'.format(join_list(column_names,
                                                                 table_name))
   heading = "REMOVE COLUMNS?"

   confirmation = yes_no_message(message, heading)
   if confirmation:

       table.Columns.Remove(column_names)

   else:
       heading = 'User chickened out'
       message = 'You decided not to delete!'
       ok_message(message, heading)


def search_columns(doc, table_name, search):

    """Removes columns from data table
    args:
    doc (Spotfire document instance): document to read from
    table_name (str): name of table
    search (str):

    """
    table = get_table(doc, table_name)
    expression = '*{}'.format(search)

    results = table.Columns.FindAll(expression)
    return results


def del_cols(doc, table_name, search):
    """Removes columns from data table
    args:
    doc (Spotfire document instance): document to read from
    table_name (str): name of table
    search (str):

    """
    delete_columns(doc, table_name, search_columns(doc, table_name, search))


def make_table(doc, table_name, column_names, data=()):
    """Makes data table in spotfire project
    args:
    doc (Spotfire document instance): document to read from
    table_name (str): name of table
    column_names (list like): names of the columns in the table
    """
    import clr
    clr.AddReference('System.Data')
    import System
    from Spotfire.Dxp.Data.Import import TextFileDataSource, TextDataReaderSettings

    from Spotfire.Dxp.Data import DataType


    col_lengths = len(column_names) -1
    if col_lengths <1:

        message = ('You are trying to create an empty table without ' +
                   'any data columns, this is useless, process stopped!')

        ok_message(message)

    else:

        text_input = [column_names]
        text_input.extend(data)
        print(text_input)
        stream = table_text_writer(text_input)

        readerSettings = TextDataReaderSettings()
        readerSettings.Separator = "\t"
        readerSettings.AddColumnNameRow(0)
        for i, col_n in enumerate(controlled_list(column_names)):

            readerSettings.SetDataType(i, DataType.String)
            print(col_n + " Added as column " + str(i))

        source = TextFileDataSource(stream, readerSettings)

        if doc.Data.Tables.Contains(table_name):
            doc.Data.Tables[table_name].ReplaceData(source)
        else:
            doc.Data.Tables.Add(table_name, source)


def add_table_column(doc, table_name, column_name):

    """Adding column to data table"""
    table = get_table(doc, table_name)


def connect_table_columns(doc, source_table_name, connect_table_name):

    """Connecting two columns from two tables"""

    pass


def copy_table(doc, table_name, newname):
    """Copies data table in spotfire project
    args:
    doc (Spotfire document instance): document to read from
    table_name (str): name of table
    """
    #from Spotfire.Dxp.Data import sourceTable
    from Spotfire.Dxp.Data.Import import DataTableDataSource
    table = get_table(doc, table_name)
    tablesource = DataTableDataSource(table)
    try:
        doc.Data.Tables.Add(newname, tablesource)
        print('Copied table %s to %s', table_name, newname)
    except ValueError:
        message = ("Table {} already exists!" +
                   " no copy will be made").format(newname)
        heading = "Warning!"
        ok_message(message, heading)

    except TypeError:

        table.ReplaceData(tablesource)
        print("This did not work")


def get_column(doc, table_name, col_name):

    """Fetches a specific column in a specific data table
    args:
       doc (Spotfire document instance): document to read from
       table_name (str): name of table to read from
       col_name (str): name of column to read

    """
    from Spotfire.Dxp.Data import DataValueCursor #, List
    table = get_table(doc, table_name)
    # place generic data cursor on a specific column
    cursor = DataValueCursor.CreateFormatted(table.Columns[col_name])

    # list object to store retrieved values
    values = []
    # iterate through table column rows to retrieve the values
    for row in table.GetRows(cursor):
	#rowIndex = row.Index ##un-comment if you want to fetch the row index
        # into some defined condition

        value = cursor.CurrentValue
        if value <> str.Empty:
            values.append(value)

    return values


def findrowfilter(table):
    from Spotfire.Dxp.Data import IndexSet
    """Returns number of rows in a table
    args:
       table (spotfire table instance): table to find rows from
    """
    #Identify the number of rows in the data table
    count = table.RowCount
    rowfilter = IndexSet(count, False)
    return rowfilter


def _delete_rows_(table, rowfilter, printstr='all rows'):

    """Deletes rows in a data table based on a filter
    args:
        table (Spotfire.table): table to delete from
        rowfilter (Spotfire.Dxp.Data.rowFilter): filter to delete with

    """
    from Spotfire.Dxp.Data import RowSelection
    heading = "REMOVE ROWS"

    message =  'Deleting {} from table {}'.format(printstr, table.Name)
    confirmation = yes_no_message(message, heading)
    if confirmation:

        table.RemoveRows(RowSelection(rowfilter))

    else:
        heading = 'User chickened out'
        message = 'You decided not to delete!'
        ok_message(message, heading)


def delete_all_rows(doc, table_name='Active'):

    """Deletes all rows in a table
    args:
      doc (Spotfire document instance): document to work with
      table_name (str): name of table to delete from
    """
    from Spotfire.Dxp.Data import IndexSet
    from Spotfire.Dxp.Data import DataValueCursor

    table = get_table(doc, table_name)
    rowfilter = findrowfilter(table)
    cursor =DataValueCursor.CreateFormatted(table.Columns[0])

    for row in table.GetRows(cursor):
        rowfilter.AddIndex(row.Index)

    rowfilter = rowfilter
    print('Rows in table' + str(len(str(rowfilter))))
    _delete_rows_(table, rowfilter)


def del_based_on_column_value(doc, column_name, remove_list, table_name='Active'):

    """Removes rows in datatable based on a list
    args:
      doc (Spotfire document instance): document to work with
      table_name (str): name of table to delete from
      column_name (str): name of column with criteria in
      remove_list (list): values of column that triggers delete
    """
    from Spotfire.Dxp.Data import DataValueCursor
    from System import Convert

    table = get_table(doc, table_name)

    cursor = []
    cursor = DataValueCursor.CreateFormatted(table.Columns[column_name])
    rowfilter = findrowfilter(table)

    for row in table.GetRows(cursor):
        if Convert.ToString(cursor.CurrentValue) in remove_list:
            rowfilter.AddIndex(row.Index)

    printstr = '{} from column {}'.format(join_list(remove_list), column_name)
    _delete_rows_(table, rowfilter, printstr)


def page_nr(doc, page_name):

    """Fetches a spotfire page, this is just a helper function because
       one has to access pages by number not name

       args:
          doc:Spotfire document instance

    returns:
        number: integer, number of page corresponding to name


    """
    logger = init_logging(__name__)
    names = get_page_names(doc)
    logger.debug('Finding page nr for %s', page_name)
    count = 0
    number = None
    for name in names:

        if name == page_name:

            number = count
            break

        count += 1

    if number is None:

        announce_no_data(page_name, names, 'Page')

    return number


def get_page(doc, page_name='Active'):

    """Fetches a page in a spotfire project
    args:
        doc:Spotfire document instance
    returns:
        page: page instance from spotfire project, or None, if page does
              not exist

    """
    if page_name == 'Active':

        page = doc.ActivePageReference

    else:
        pnr = page_nr(doc, page_name)

        if pnr is not None:

            page = doc.Pages[pnr]

        else:

            page = None

    return page


def get_panel(doc, page_name, panel_type):
   """Fetches a panel on a page
   args:
        doc:Spotfire document instance
        page_name (str): name of page
        panel_type (str): name of panel
   returns: ret_panel: spotfire.Panel
   """

   ids = {'filter': 'TypeIdentifier:Spotfire.FilterPanel',
          'details': 'TypeIdentifier.Spotfire.DetailsOnDemandPanel',
          'data': 'TypeIdentifier.Spotfire.DataPanel'}
   ret_panel = None
   if panel_type not in ids.keys():

        message = ('{} is not a valid panel type, please ' +
                   'choose among [{}]').format(panel_type,
                                               join_list(sorted(ids)))
        ok_message(message)
   else:
       page = get_page(doc, page_name)
       ret_panel = None
       for panel in page.Panels:
           if str(panel.TypeId) == ids[panel_type]:

               ret_panel = panel

   return ret_panel


def get_panel_group(doc, page_name, panel_type, group_name):

    """Fetches a panel group on a page

    args:
        doc:Spotfire document instance
        page_name (str): name of page
        panel_type (str): name of panel
        group_name (str): name of panel group
   returns: ret_group: spotfire.Panel
    """
    panel = get_panel(doc, page_name, panel_type)

    ret_group = None
    for group in panel.TableGroups:

        if group.Name == group_name:

            ret_group = group

    return ret_group


def get_group_filter(doc, page_name, group_name, filter_name):

    """Fetches a panel filter group on a page

    args:
        doc:Spotfire document instance
        page_name (str): name of page
        group_name (str): name of panel group
        filter_name (str): name of filter
    returns: ret_filter: spotfire.Panel
    """
    group = get_panel_group(doc, page_name, 'filter', group_name)
    ret_filter = group.GetFilter(filter_name)
    if ret_filter is None:

        message = ('No filter with name "{}" in group "{}" on ' +
                   'page "{}"').format(filter_name, group_name, page_name)
        ok_message(message)

    return ret_filter


def get_filter_selection(doc, page_name, group_name, filter_name):
    """Getting selected items in a filter on a filter panel
    args:
        doc:Spotfire document instance
        page_name (str): name of page
        group_name (str): name of panel group
        filter_name (str): name of filter

    returns: item_list (list): list of items from filter
    """
    select_filter = get_group_filter(doc, page_name, group_name, filter_name)
    item_list = []
    checkbox_filter = select_filter.As [CheckBoxFilter]()
    for item in checkbox_filter:
        if checkbox_filter.IsChecked(item):

            item_list.append(item)

    return item_list

