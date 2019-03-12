"""Author dbs: helper functions for use in spotfire scripts"""
import logging
import sys
import re

def init_logging(name,loglevel=False,file_name=None,full=False):

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

       format='%(asctime)s: %(name)s - %(levelname)s - %(message)s'

    else:

       format='%(asctime)s: %(message)s'

    format='%(asctime)s: %(message)s'
    dateformat='%m/%d/%Y %I:%M:%S'
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

      logging.basicConfig(level=numeric_level,format=format,datefmt=dateformat)
      logger = logging.getLogger(__name__)

      #Adds file handle to logger if file_name is specified
      if file_name is not None:

         try:
            fh = logging.FileHandler(file_name)
            formatter = logging.Formatter(format)
            fh.setFormatter(formatter)
            logger.addHandler(fh)

         except IOError as io:

            raise io

    return(logger)


def announce_no_data(name,valid_names,data_type):

    """Announcing that no data is found in a message box

       args:

           name     : str
           valid_data: list
           data_type :  str, name of data type not found

       raises: AttributeError if data_type is not in valid_types
    """

    logger=init_logging(__name__)
    valid_types=['Page','Table','Chart','Column']

    if data_type not in valid_types:

        raise AttributeError('No such data type in spotfire project')

    else:
        heading='Page does not exist!'
        message=('{} with name "{}" does not exist\n'.format(data_type,name)+
                 'Please choose among the following: \n{}'.format(
                 '\n'.join(valid_names)))
        ok_message(message,heading)
        logger.warning(message)
        quit


def ok_message(message,heading):

    """ Creates information box
        args:

            message: str, what is what is displayed in box

            heading: str, what is displayed as heading 'name' of box

    """
    logger=init_logging(__name__)
    import clr
    clr.AddReference('System.Windows.Forms')
    from System.Windows import Forms
    Forms.MessageBox.Show(message,heading,Forms.MessageBoxButtons.OK)


def yes_no_message(message,heading):

    """ Creates decision box with yes and no, returns boolean

        args:

            message: str, what is what is displayed in box

            heading: str, what is displayed as heading 'name' of box

            returns:
                  yes: boolean

    """
    logger=init_logging(__name__)
    import clr
    clr.AddReference('System.Windows.Forms')
    from System.Windows import Forms
    yes=True
    reply=Forms.MessageBox.Show(message,heading,Forms.MessageBoxButtons.YesNo)
    if reply==Forms.DialogResult.No:

        yes=False
        logger.debug('User pressed no')

    else:
        logger.debug('User pressed yes')

    return(yes)


def get_page_names(doc):

    """Fetches names of the different pages in a spotfire project

    args:
        doc:Spotfire document instance

    returns:
        names: list of strings, names of the different pages

    """
    logger=init_logging(__name__)
    names=[]
    logger.debug('Looping through pages')
    for page in doc.Pages:

	names.append(page.Title)

    logger.debug('Returning table')
    logger.debug(names)

    return(names)


def get_table_names(doc):

    """Fetches the data table names in a spotfire project

    args:
        doc:Spotfire document instance

    returns:
        names: list of strings, names of the different tables

    """
    logger=init_logging(__name__)
    names=[]
    logger.debug('Looping through tables')
    for table in doc.Data.Tables:

        names.append(table.Name)

    logger.debug('Returning table')
    return(names)


def get_visual_names(page,viz_type='Chart'):
    """Fetches the visualization names on a page in a spotfire project

    args:
        page:Spotfire document page instance

    returns:
        visuals: dictionary of strings, names of the different visualizations


    """
    import Spotfire.Dxp.Application.Visuals as viz
    visuals={}
    for visual in page.Visuals:
        name=visual.Title
        viz_id=visual.TypeId.Name.replace('Spotfire.','')
        if re.match(r'.*{}'.format(viz_type),viz_id):

            print(name)
            print(viz_id)
            visuals[name]=viz_id

    return(visuals)


def get_chart_names(page):

    """Fetches the chart names on a page in a spotfire project

    args:

       page: Spotfire document page instance

       returns:

           charts: dictionary with keys and values as strings

    """

    charts=get_visual_names(page)

    return(charts)


def get_table(doc,name):

    """Fetches a specific datatable
    args:
        doc: Spotfire document instance

    returns:
        table: Spotfire data table

    raises: AttributeError: if table with name does not exist

    """
    logger=init_logging(__name__)
    (tabletest,table)=doc.Data.Tables.TryGetValue(name)
    if tabletest:

        return(table)

    else:

        valid_tables=get_table_names(doc)
        announce_no_data(name,valid_tables,'Table')


def page_nr(doc,page_name):

    """Fetches a spotfire page, this is just a helper function because
       one has to access pages by number not name

       args:
          doc:Spotfire document instance

    returns:
        nr: integer, number of page corresponding to name


    """
    logger=init_logging(__name__)
    names=get_page_names(doc)
    count=0
    nr=None
    for name in names:

        if name==page_name:

            nr=count
            break

        count+=1

    if nr is None:

        announce_no_data(page_name,names,'Page')

    return(nr)


def get_page(doc,page_name):

    """Fetches a page in a spotfire project
    args:
        doc:Spotfire document instance
    returns:
        page: page instance from spotfire project, or None, if page does
              not exist

    """
    pnr=page_nr(doc,page_name)

    if pnr is not None:

        page=doc.Pages[pnr]

    else:

        page=None

    return(page)
