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


def add_page(doc, page_name, ask=False):
    """Add page in document
    args:
    doc (Spotfire document instance): document to read from
    page_name (str): name of page to add
    """
    pages = doc_pages(doc)
    page_names = get_page_names(doc)
    confirmation= True
    if page_name in page_names and ask:
        heading = 'Overwrite page'
        message = ('You are trying to create a page, but the page name "{}"' +
                   ' exists, do you want to overwrite it? Answering no will' +
                   ' keep the page, yes will overwrite it').format(page_name)
        confirmation = yes_no_message(message, heading)
        if confirmation:
          delete_page(doc, page_name)

    if confirmation:
        pages.AddNew(page_name)


def get_page(doc, page_name='Active', ask_create=False):

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
        for check_page in pages:
            name = check_page.Title
            LOGGER.debug('Page name: %s', name)
            if name == page_name:
                page = check_page
                break
    LOGGER.debug('Here is page so far:')
    LOGGER.debug(page)
    confirmation = True

    if page is None:
        LOGGER.debug(('After looking through pages, page with name:' +
                      ' %s was not found'), page_name)
        confirmation = False
        if ask_create:
            LOGGER.debug('Will ask to create page %s', page_name)
            heading = 'Page does not exist!'
            message = 'Do you wanna create page {} ?'.format(page_name)
            confirmation = yes_no_message(message, heading)
        else:
            confirmation = True

    LOGGER.debug('Confirmation is:')
    LOGGER.debug(confirmation)
    if confirmation:
        add_page(doc, page_name, False)
        page = get_page(doc, page_name)

    return page
