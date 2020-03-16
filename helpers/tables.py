import re
from basic import init_logging, list_string, controlled_list, create_line
from basic import get_list_entry
from boxes import deletion_message, yes_no_message, ok_message
from boxes import announce_no_data

LOGGER = init_logging(__name__)

def doc_tables(doc):
    """Gets tables from a spotfire document
    args:
    doc (Spotfire document instance): document to read from
    """
    LOGGER.debug('Passed data type')
    LOGGER.debug(type(doc))
    return doc.Data.Tables


def get_table_names(doc):

    """Fetches the data table names in a spotfire project
    args:
       doc (Spotfire document instance): document to find tables from
    returns:
        names: list of strings, names of the different tables

    """
    LOGGER = init_logging(__name__)
    names = []
    LOGGER.debug('Looping through tables')
    for table in doc_tables(doc):

        names.append(table.Name)

    LOGGER.debug('Returning table')
    return names


def delete_tables(doc, table_names):
    """Deletes tables in document from a list
    args:
    doc (Spotfire document instance): document to read from
    table_names (list of strings): names of tables to delete
    """
    LOGGER.debug('About to delete tables')
    LOGGER.debug(table_names)
    if table_names:
        if type(table_names) == str:
            table_list = []
            table_list.append(table_names)
            table_names = table_list
        LOGGER.debug('Passed conversion')
        LOGGER.debug(table_names)
        confirmation = deletion_message(table_names, 'table')

        if confirmation:
            tables = doc_tables(doc)
            for table_name in table_names:
                table = tables[table_name]
                tables.Remove(table)
                LOGGER.debug('%s deleted'.format(table_name))

        else:
           no_deletion(table_names, 'table')
    else:
        LOGGER.debug('No deletion, empty list supplied')


def delete_table(doc, table_name):
    """Deletes table in document
    args:
    doc (Spotfire document instance): document to read from
    table_name (str): name of table to delete
    """
    delete_tables(doc, table_name)


def delete_all_tables(doc):
    """Deletes all table in document
    args:
    doc (Spotfire document instance): document to read from
    """
    table_names = get_table_names(doc)
    delete_tables(doc, table_names)


def delete_tables_search(doc, pattern, regex=False):
    """Deletes table in document with pattern in name
    args:
    doc (Spotfire document instance): document to read from
    pattern (string): pattern to find in table_name
    """
    table_names = get_table_names(doc)
    del_table_names = []
    for table_name in table_names:
        found = False
        if regex:
            match =re.match(r'{}'.format(pattern), table_name)
            if match:
                found = True
        else:
            if pattern in table_name:
                found = True

        if found:
            del_table_names.append(table_name)
    LOGGER.debug('Tables matching  pattern %s: %s',
                 pattern, list_string(del_table_names))

    delete_tables(doc, del_table_names)


def get_table(doc, table_name='Active'):

    """Fetches a specific datatable
    args:
      doc (Spotfire document instance): Document to work with
      table_name (str): name of table

    returns:
        table: None or Spotfire data table

    raises: AttributeError: if table with name does not exist

    """
    LOGGER.debug('Fetching table with name %s', table_name)
    table = None
    if table_name == 'Active':

        table = doc.ActiveDataTableReference

    else:
        tables = doc_tables(doc)
        try:
            table = tables[table_name]
            # doc.Data.Tables[table_name]
        except IndexError:

            valid_tables = get_table_names(doc)
            announce_no_data(name, valid_tables, 'table')

    return table


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
        print('Copied table {} to {}'.format(table_name, newname))
    except ValueError:
        message = ("Table {} already exists!" +
                   " no copy will be made").format(newname)
        heading = "Warning!"
        ok_message(message, heading)

    except TypeError:

        table.ReplaceData(tablesource)
        print("This did not work")


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

    tables = doc_tables(doc)
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

        if tables.Contains(table_name):
            tables[table_name].ReplaceData(source)
        else:
            tables.Add(table_name, source)


def make_table_from_csv(doc,table_name, file_path):
    """Makes table from csv
    args:
    doc (Spotfire document instance): document to read from
    table_name (str): name of table
    """
    from Spotfire.Dxp.Data.Import import TextFileDataSource, TextDataReaderSettings

    file_path=r'{}'.format(file_path)
    # specify any settings for the file
    settings= TextDataReaderSettings()
    settings.Separator=','
    dataSource=TextFileDataSource(file_path, settings)

    tables = doc_tables(doc)
    #datasource name and datasource to add table
    tables.Add(table_name, dataSource)


def define_relation(doc, first_table_name, second_table_name,
                    first_col_name, second_col_name=None, box_message=True):
    """Sets relation between table columns in spotfire project
    args:
    doc (Spotfire document instance): document to read from
    first_table_name (str): name of table
    second_table_name (str): name of table
    first_col_name (str): name of first column
    second_col_name (str or None): name of second column, is
       defaulted to None, if None, then it is assumed that the
       column name is common between the tables, and equal to
       first_col_name
    box_message (bool): decides if user will get message in pop up box
    """
    from Spotfire.Dxp.Data import DataRelation
    if second_col_name is None:
        second_col_name = first_col_name
    tables = doc_tables(doc)
    # set the left and right tables
    try:
        first= tables[first_table_name]
        second= tables[second_table_name]
        # add the relation
        # "Region" is the column that relates the two tables
        relation_string = "[{}].[{}] = [{}].[{}]".format(first_table_name,
                                                         first_col_name,
                                                         second_table_name,
                                                         second_col_name)
        doc.Data.Relations.Add(first, second, relation_string)

    except KeyError:
        message = ('Could not set up relation between {} in table {}\n' +
                   ' and {} in table {}\n.' +
                   'Ensure that both tables exist').format(first_table_name,
                                                         first_col_name,
                                                         second_table_name,
                                                         second_col_name)
        if box_message:
            ok_message(message)


def get_vector_names(doc, table_name):
    """Fetches vector names from columns in a table
    args:
    doc (Spotfire document instance): document to read from
    table_name (str): name of table
    returns. vector_names (list): names of vectors
    """
    col_names = get_column_names(doc, table_name)
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


def get_well_names(doc, table_name):

    """Fetches well_names from columns in a table
    args:
    doc (Spotfire document instance): document to read from
    table_name (str): name of table
    returns. well_names (list): names of wells
    """
    LOGGER.debug('Fetching table with name %s', table_name)
    col_names = get_column_names(doc, table_name)
    well_names = []
    for col_name in col_names:
        m = re.match(r'W[^:]+.(.*)', col_name)
        if m:
            well_name = m.group(1)
            if well_name not in well_names:
                well_names.append(well_name)
        else:
            LOGGER.debug(col_name + " does not match")

    return well_names


def make_vector_and_well_tables(doc, in_table_name):

    """Makes tables with names of vectors and wells
    args
    doc (Spotfire document instance): document to read from
    in_table_name (str): name of table to extract from

    """
    vector_names = get_vector_names(doc, in_table_name)
    well_names = get_well_names(doc, in_table_name)

    vector_table_name = 'Vectors'
    well_table_name = 'Wells'
    if in_table_name in [vector_table_name, well_table_name]:
        message = ('The name of the table to create well and vector names' +
                   ' from corresponds to the name of one of the tables.\n ' +
                   'This is not a particularly clever idea, so try again with\n'+
                   ' a different name!')
        ok_message(message)
    else:
        make_table(doc, vector_table_name, vector_table_name , vector_names)
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


def add_relations(doc):

    """Adds relations between columns with common name across all tables
    args:
    doc (Spotfire document instance): document to read from
    """
    table_names = get_table_names(doc)
    name_count = {}
    # Finding common columns
    for table_name in table_names:
        column_names = get_column_names(doc, table_name)
        for column_name in column_names:
            name_count[column_name] = name_count.get(column_name,0) + 1
    relation_names = []
    for col_name in name_count:
        if name_count[col_name] > 1:
            relation_names.append(col_name)

    base_table_name = table_names.pop(0)

    for table_name in table_names:
        for col_name in relation_names:
            define_relation(doc, base_table_name, table_name, col_name)


def get_column_names(doc, table_name):

    """Fetches the column names from a data table
    args:
    doc (Spotfire document instance): document to read from
    table_name (str): name of table
    returns. col_names (list): names of the columns
    """
    table = get_table(doc, table_name)
    col_names = [c.Name for c in table.Columns]
    return col_names


def delete_columns(doc, table_name, column_names, box_message=True):
   """Removes columns from data table
    args:
    doc (Spotfire document instance): document to read from
    table_name (str): name of table
    col_list (list): names of the columns to be deleted
    """
   table = get_table(doc, table_name)
   extra_text = 'in table {}'.format(table_name)
   if column_names:
       if type(column_names) == str:
           column_list = []
           column_list.append(column_names)
           column_names = column_list

   confirmation = deletion_message(column_names, 'column',extra_text)
   if confirmation:
       for column_name in column_names:
           try:
               table.Columns.Remove(column_name)
           except ValueError:
               message = 'No column named {} in table {}'.format(column_name,
                                                                 table_name)
               if box_message:
                   ok_message(message)
               else:
                   LOGGER.warning(message)

   else:
       no_deletion(column_names, extra_text)


def delete_column(doc, table_name, column_name):
    """Removes column from data table
    args:
    doc (Spotfire document instance): document to read from
    table_name (str): name of table
    """
    delete_columns(doc, table_name, column_name)


def delete_columns_search(doc, table_name, pattern, regex=False):
    """Deletes page in document with pattern in name
    args:
    doc (Spotfire document instance): document to read from
    pattern (string): pattern to find in page_name
    """
    column_names = get_column_names(doc, table_name)
    del_column_names = []
    LOGGER.debug('Searching for pattern %s', pattern)
    for column_name in column_names:
        LOGGER.debug(column_name)
        found = False
        if regex:
            match =re.match(r'{}'.format(pattern), column_name)

            if match:
                found = True
        else:
            if pattern in column_name:
                found = True

        if found:
            LOGGER.debug('This is a match %s', column_name)
            del_column_names.append(column_name)
    LOGGER.debug('Columns matching  pattern %s: %s',
                 pattern, list_string(del_column_names))

    delete_columns(doc, table_name, del_column_names)


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


def get_column_as_list(doc, table_name, col_name):

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
