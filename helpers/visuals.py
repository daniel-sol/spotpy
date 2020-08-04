"""Author dbs: helper functions for use in spotfire scripts"""
import logging
import sys
import re
from . import basic
from . import pages
from . import tables


LOGGER = basic.init_logging(__name__)


def viz_as_content(viz):
    """ Returns a vizualization as visual content
    args:
      viz (Spotfire.Dxp.Application.Visuals.Visual): visual
    returns viz_content (Spotfire.Dxp.Application.Visuals.VisualContent)
    """

    from Spotfire.Dxp.Application.Visuals import VisualContent
    viz_content = viz.As[VisualContent]()
    return viz_content


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
    page = pages.get_page(doc, page_name)
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

        page = pages.get_page(doc, page_name)
        viz = get_visual(doc, page_name, viz_name)

        if viz is None:
            viz = page.Visuals.AddNew[viz_dict[viz_type]]()
            message = 'Making visual {} on page {}'.format(viz_name,
                                                           page_name)
            heading = "INFO"
            ok_message(message, heading)

        print(viz)
        viz.Data.DataTableReference = tables.get_page(doc, table_name)
        try:
            viz_cont = viz_as_content(viz)
            viz_cont.Title = viz_name
            viz_cont.ShowTitle = True
        except AttributeError:
            print('Cannot set title')

        print('Created {}'.format(viz_name))
    return viz


def make_linechart(doc, table_name, x_name, y_names,
                   y_stat_names, line_by_names, page_name='Active',
                   viz_name=None):

    """Adds vizualisation to page
    args:
    doc (Spotfire document instance): document to read from
    page_name (str): name of page
    viz_name (str): name of visual
    returns viz (spotfire visual)
    """
    viz_type = 'linechart'
    viz = add_viz_to_page(doc, viz_type, viz_name, table_name, page_name)
    viz_content = viz_as_content(viz)
    if viz_name is not None:
        viz_content.Title = viz_name

    if x_name == 'DATE':

        x_expression = 'BinByDateTime([DATE],"Year.Month",2)'

    else:
        x_expression = "[" + x_name + "]"
    viz.XAxis.Expression = x_expression

    # single line
    color_expression = '<[Axis.Default.Names]>'

    viz.YAxis.Expression = basic.create_expression(y_names, y_stat_names)
    viz.LineByAxis.Expression = basic.create_expression(line_by_names,'')
    viz.ColorAxis.Expression = color_expression

    viz.Legend.Visible = True


def get_html(doc, page_name):

    from Spotfire.Dxp.Application.Visuals import HtmlTextArea
    viz = get_visual(doc, page_name, 'Text Area')


    html = viz.As[HtmlTextArea]().HtmlContent
    return html


def make_histograms(doc, table_name, column_names, page_name, color_by, bins):
    """Adds histograms to page
    args:
    doc (Spotfire document instance): document to read from
    table_name (str): name of table to plot data from
    column_names (list of str): names of columns to plot
    page_name (str): name of page
    color_byt (list of str): names of columns to colour by
    bins (int): number of bins
    """

    table = tables.get_table(doc, table_name)
    pages.get_page(doc, page_name, True)
    viz_type = 'barchart'
    for colum_name in column_names:

        add_viz_to_page(doc, viz_type, column_name, table_name, page_name)



def make_histograms_search(doc, table_name, column_search, page_name, color_by,
                           bins):

    column_names =  tables.get_column_names_search(doc, table_name,
                                                   column_search,
                                                   regex=True)

    LOGGER.debug('Found %s', basic.list_string(column_names))

    make_histograms(doc, table_name, column_names, page_name, color_by, bins)


