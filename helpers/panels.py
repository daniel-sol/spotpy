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

