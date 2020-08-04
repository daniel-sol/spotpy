from basic import init_logging
from boxes import ok_message, yes_no_message, announce_no_data

LOGGER = init_logging(__name__)

def get_coloraxis(viz_cont):
    """Gets color axis from visual_content
       args:  viz_content: Spotfire.Dxp.Application.Visuals.VisualContent
       returns col_axis:  Spotfire.Dxp.Application.Visuals.ColorAxis
    """
    col_axis = None
    if hassatr(viz_cont, 'ColorAxis'):
        col_axis = viz_cont.ColorAxis
    else:
        message = 'Visual has no color axis'
        ok_message(message)

    return col_axis


def make_categorical_rule(viz_content):
    """ Sets color axis to categorical for chart
       args:  viz_content: Spotfire.Dxp.Application.Visuals.VisualContent
    """
    col_axis = get_coloraxis(viz_content)
    color_rule = col_axis.coloring.AddCategoricalColorRule()
    return color_rule


def make_continuous_rule(viz_content):
    """ Sets color axis to categorical for chart
       args:  viz_content: Spotfire.Dxp.Application.Visuals.VisualContent
    """
    col_axis = get_coloraxis(viz_content)
    color_rule = col_axis.coloring.AddContinuousColorRule()
    return color_rule


def set_colour_rule(viz_cont, expression):
    """ Sets colour rule for vizualization
    args:
    doc (Spotfire document instance): document to read from
    """
    from Spotfire.Dxp.Application.Visuals import ColorAxis, VisualContent
    from System.Drawing import Color
    v = v.As[VisualContent]()
    if hasattr(v,'ColorAxis'):
        v.ColorAxis.Expression=expr
        c= v.ColorAxis.Coloring
        c.AddCategoricalColorRule()


