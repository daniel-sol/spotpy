import re
from basic import init_logging, list_string
from boxes import deletion_message, yes_no_message, announce_no_data

LOGGER = init_logging(__name__)


def doc_pages(doc):
    """Gets pages from a spotfire document
    args:
    doc (Spotfire document instance): document to read from
    """
    return doc.Pages


def get_page_names(doc):

    """Fetches names of the different pages in a spotfire project
    args:
    doc (Spotfire document instance): document to read from
    returns:
        names: list of strings, names of the different pages

    """

    names = []
    LOGGER.debug('Looping through pages')
    for page in doc_pages(doc):
        page_name = page.Title
        names.append(page_name)

        LOGGER.debug('Returning table name %s', page_name)
    LOGGER.debug(names)

    return names


def delete_pages(doc, page_names):
    """Deletes pages in document from a list
    args:
    doc (Spotfire document instance): document to read from
    page_names (list of strings): names of pages to delete
    """
    LOGGER.debug('About to delete pages')
    LOGGER.debug(page_names)
    if page_names:
        if type(page_names) == str:
            page_list = []
            page_list.append(page_names)
            page_names = page_list
        LOGGER.debug('Passed conversion')
        LOGGER.debug(page_names)
        confirmation = deletion_message(page_names, 'page')

        if confirmation:
            pages = doc_pages(doc)
            for page in doc_pages(doc):
                if page.Title in page_names:
                    pages.Remove(page)
                    LOGGER.debug('%s deleted'.format(page.Title))

        else:
           no_deletion(page_names, 'page')
    else:
        LOGGER.debug('No deletion, empty list supplied')


def delete_page(doc, page_name):
    """Deletes page in document
    args:
    doc (Spotfire document instance): document to read from
    page_name (str): name of page to delete
    """

    delete_pages(doc, page_name)


def delete_all_pages(doc):
    """Deletes all page in document
    args:
    doc (Spotfire document instance): document to read from
    """
    page_names = get_page_names(doc)
    delete_pages(doc, page_names)


def delete_pages_search(doc, pattern, regex=False):
    """Deletes page in document with pattern in name
    args:
    doc (Spotfire document instance): document to read from
    pattern (string): pattern to find in page_name
    """
    page_names = get_page_names(doc)
    del_page_names = []
    for page_name in page_names:
        found = False
        if regex:
            match =re.match(r'{}'.format(pattern), page_name)
            if match:
                found = True
        else:
            if pattern in page_name:
                found = True

        if found:
            del_page_names.append(page_name)
    LOGGER.debug('Pages matching  pattern %s: %s',
                 pattern, list_string(del_page_names))

    delete_pages(doc, del_page_names)


def add_page(doc, page_name):
    """Add page in document
    args:
    doc (Spotfire document instance): document to read from
    page_name (str): name of page to add
    """
    pages = doc_pages(doc)
    page_names = get_page_names(doc)
    confirmation= True
    if page_name in page_names:
        heading = 'Overwrite page'
        message = ('You are trying to create a page, but the page name "{}"' +
                   ' exists, do you want to overwrite it? Answering no will' +
                   ' keep the page, yes will overwrite it').format(page_name)
        confirmation = yes_no_message(message, heading)
        if confirmation:
          delete_page(doc, page_name)

    if confirmation:
        pages.AddNew(page_name)


def get_page(doc, page_name='Active'):

    """Fetches a page in a spotfire project
    args:
        doc:Spotfire document instance
    returns:
        page: page instance from spotfire project, or None, if page does
              not exist

    """
    pages = doc_pages(doc)
    page = None
    if page_name == 'Active':

        page = doc.ActivePageReference

    else:
        for page in pages:
            if page.Title == page_name:
                break
    if page is None:

        page_names = get_page_names(doc)
        announce_no_data(page_name, page_names, 'page')
        pass

    return page
