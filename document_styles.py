"""
Document styling utilities for quote generation.

This module provides functions for styling cells, borders,
and backgrounds in Word documents.
"""

from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def set_cell_border(cell, **kwargs):
    """
    Set cell borders.

    Args:
        cell: Table cell to style
        **kwargs: Border specifications for each edge
            (top, left, bottom, right)

    Usage:
        set_cell_border(
            cell,
            top={"sz": 12, "val": "single", "color": "#000000"}
        )
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    edges = ('top', 'left', 'bottom', 'right')
    for edge in edges:
        if edge in kwargs:
            edge_data = kwargs.get(edge)
            edge_el = OxmlElement(f'w:{edge}')
            for key, value in edge_data.items():
                edge_el.set(qn(f'w:{key}'), str(value))
            tcPr.append(edge_el)


def set_cell_background(cell, fill):
    """
    Set cell background color.

    Args:
        cell: Table cell to style
        fill: Hex color code (e.g., '000000' for black)
    """
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), fill)
    cell._element.get_or_add_tcPr().append(shading_elm)
