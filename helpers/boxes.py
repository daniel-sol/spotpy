from basic import init_logging, list_string, item_dict


LOGGER = init_logging(__name__)


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
    import clr
    clr.AddReference('System.Windows.Forms')
    from System.Windows import Forms
    answer = False
    reply = Forms.MessageBox.Show(message, heading, Forms.MessageBoxButtons.YesNo)
    if reply == Forms.DialogResult.No:

        LOGGER.debug('User pressed no')

    else:
        answer = True
        LOGGER.debug('User pressed yes')

    return answer


def find_item(item_type):
    """returns string to confirm items"""
    items = item_dict()
    LOGGER.debug(items)
    LOGGER.debug('Extracting with |%s|', item_type)
    item_types = None
    try:
        item_types = items[item_type]
    except KeyError:
        LOGGER.warning('This was not right')
        raise KeyError('%s is not in {}'.format(item_type,
                                                list_string(items)))
    return item_types


def deletion_message(del_list, item_type, extra_text=''):
    """Asks for permision to delete elements
    args:
    del_list (list): list of items to delete
    item_type (str): type to delete
    extra_text (str): extra text to add
    """
    LOGGER.debug('Deletion of type %s', item_type)
    del_type = find_item(item_type)
    LOGGER.debug('Passing on %s', del_type)
    heading = 'Deletion of {}'.format(del_type)
    message = ('You are about to delete the following items \n{}\n' +
               ' {} do you want to proceed?').format(list_string(del_list),
                                                     extra_text)

    confirmation = yes_no_message(message, heading)
    LOGGER.debug('Received answer')
    LOGGER.debug(confirmation)
    return confirmation


def no_deletion(del_list, item_type, extra_text=''):
    """ Confirms no deletion of items
    args:
    del_list (list): list of items to delete
    item_type (str): type to delete
    extra_text (str): extra text to add
    """
    heading = 'No deletion'
    message = ('You decided not to delete {} {}' +
               'from {}').format(list_string(del_list), item_type, extra_text)
    ok_message(message, heading)


def announce_no_data(name, valid_names, data_type):

    """Announcing that no data is found in a message box

       args:
           name     : str
           valid_data: list
           data_type :  str, name of data type not found

       raises: AttributeError if data_type is not in valid_types
    """


    valid_types = sorted(item_dict())
    LOGGER.debug(valid_types)
    if data_type not in valid_types:

        raise AttributeError('No such data type in spotfire project')

    else:
        heading = 'Data does not exist!'
        message = ('{} with name {} does not exist,\n please choose among' +
                   ' the following {}').format(data_type,
                                        name, list_string(valid_names))

        ok_message(message, heading)
        LOGGER.debug('Written message')
