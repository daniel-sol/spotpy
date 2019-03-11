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

        #raise AttributeError('Page with name {} does not exist'.format(page_name))
        logger.info('This is not a valid page name')
        logger.info('Please choose among the following')
        logger.info(names)
        exit()

    return(nr)


def get_page(doc,page_name):
    """Fetches a a page in a spotfire project
    args:
        doc:Spotfire document instance
    returns:
        page: page instance from spotfire project

    raises: AttributeError: if page with name page_name does not exist
    """
    pnr=page_nr(doc,page_name)

    if pnr is not None:

        page=doc.Pages[pnr]

    else:
        page=None

    return(page)

def get_visual_names(page):
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
        if re.match(r'.*Chart',viz_id):

            print(name)
            print(viz_id)
            visuals[name]=viz_id

    return(visuals)


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
        logger.info('This is not a valid table name')
        logger.info('Please choose among the following')
        logger.info(valid_tables)
        raise AttributeError('Table {} does not exist'.format(name))





