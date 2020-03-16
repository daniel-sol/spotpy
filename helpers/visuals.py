"""Author dbs: helper functions for use in spotfire scripts"""
import logging
import sys
import re


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


def make_linechart(doc, viz_type, viz_name, table_name, x_name, y_names,
                   line_by_names, page_name='Active'):

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
    # single line
    color_expression = '<[Axis.Default.Names]>'

    if not isinstance(y_names, list):

        y_expression = y_expression + 'Avg (' + y_names + ')'

    else:

        y_expression = ' Avg (' + y_names.pop(0) + ')'

        for y_name in y_names:
            y_expression = y_expression + ', Avg (' + y_name + ')'
    viz.LineByAxis.Expression = '<[Type] NEST [Input Type]>'
    viz.YAxis.Expression = y_expression
    viz.ColorAxis.Expression = color_expression

    viz.Legend.Visible = True


def get_html(doc, page_name):

    from Spotfire.Dxp.Application.Visuals import HtmlTextArea
    viz = get_visual(doc, page_name, 'Text Area')


    html = viz.As[HtmlTextArea]().HtmlContent
    return html
