# coding=utf-8

from collections import namedtuple

from pyrevit import revit, DB
from pyrevit.compat import safe_strtype
from pyrevit import forms
from pyrevit import coreutils
from pyrevit import EXEC_PARAMS
from pyrevit import HOST_APP
from pyrevit.coreutils import pyutils

from dosymep_libs.bim4everyone import *

logger = script.get_logger()
output = script.get_output()

ParamDef = namedtuple('ParamDef', ['name', 'type'])


def is_calculable_param(param):
    if param.StorageType == DB.StorageType.Double:
        return True

    if param.StorageType == DB.StorageType.Integer:
        val_str = param.AsValueString()
        if val_str and safe_strtype(val_str).lower().isdigit():
            return True

    return False


def calc_param_total(element_list, param_name):
    sum_total = 0.0

    def _add_total(total, param):
        if param.StorageType == DB.StorageType.Double:
            total += param.AsDouble()
        elif param.StorageType == DB.StorageType.Integer:
            total += param.AsInteger()

        return total

    for el in element_list:
        param = el.LookupParameter(param_name)
        if not param:
            el_type = revit.doc.GetElement(el.GetTypeId())
            type_param = el_type.LookupParameter(param_name)
            if not type_param:
                logger.error('Element with ID: {} '
                             'does not have parameter: {}.'.format(el.Id,
                                                                   param_name))
            else:
                sum_total = _add_total(sum_total, type_param)
        else:
            sum_total = _add_total(sum_total, param)

    return sum_total


def format_length(total):
    return '{} метров\n' \
           '{} сантиметров'.format(total / 3.28084,
                                   (total / 3.28084) * 100)


def format_area(total):
    return '{} квадратных метров\n' \
           '{} квадратных сантиметров'.format(total / 10.7639,
                                              (total / 10.7639) * 10000)


def format_volume(total):
    return '{} кубических метров\n' \
           '{} кубических сантиметров'.format(total / 35.3147,
                                              (total / 35.3147) * 1000000)


if HOST_APP.is_newer_than(2022, or_equal=True):
    formatter_funcs = {DB.SpecTypeId.Length: format_length,
                       DB.SpecTypeId.Area: format_area,
                       DB.SpecTypeId.Volume: format_volume}
else:
    formatter_funcs = {DB.ParameterType.Length: format_length,
                       DB.ParameterType.Area: format_area,
                       DB.ParameterType.Volume: format_volume}


def output_param_total(element_list, param_def):
    total_value = calc_param_total(element_list, param_def.name)

    print('Итого параметр \"{}\":\n\n'.format(param_def.name))
    if param_def.type in formatter_funcs.keys():
        outputstr = formatter_funcs[param_def.type](total_value)
    else:
        outputstr = '{}\n'.format(total_value)
    print(outputstr)


def output_breakdown(element_list, param_def):
    for element in element_list:
        total_value = calc_param_total([element], param_def.name)

        if param_def.type in formatter_funcs.keys():
            outputstr = formatter_funcs[param_def.type](total_value)
        else:
            outputstr = '{}\n'.format(total_value)
        print('{}\n{}'.format(output.linkify(element.Id), outputstr))


def process_options(element_list):
    # find all relevant parameters
    param_sets = []

    for el in element_list:
        shared_params = set()
        # find element parameters
        for param in el.ParametersMap:
            if is_calculable_param(param):
                pdef = param.Definition
                if HOST_APP.is_newer_than(2022, or_equal=True):
                    shared_params.add(ParamDef(pdef.Name,
                                               pdef.GetDataType()))
                else:
                    shared_params.add(ParamDef(pdef.Name,
                                               pdef.ParameterType))

        # find element type parameters
        el_type = revit.doc.GetElement(el.GetTypeId())
        if el_type and el_type.Id != DB.ElementId.InvalidElementId:
            for type_param in el_type.ParametersMap:
                if is_calculable_param(type_param):
                    pdef = type_param.Definition
                    if HOST_APP.is_newer_than(2022, or_equal=True):
                        shared_params.add(ParamDef(pdef.Name,
                                                   pdef.GetDataType()))
                    else:
                        shared_params.add(ParamDef(pdef.Name,
                                                   pdef.ParameterType))

        param_sets.append(shared_params)

    # make a list of options from discovered parameters
    if param_sets:
        all_shared_params = param_sets[0]
        for param_set in param_sets[1:]:
            all_shared_params = all_shared_params.intersection(param_set)

        if HOST_APP.is_newer_than(2022, or_equal=True):
            return {'{} <{}>'.format(x.name, str(x.type.TypeId).split(":")[-1].split("-")[0].capitalize()): x
                    for x in all_shared_params}
        else:
            return {'{} <{}>'.format(x.name, x.type): x
                    for x in all_shared_params}


def process_sets(element_list):
    el_sets = pyutils.DefaultOrderedDict(list)

    # add all elements as first set, for totals of all elements
    el_sets['Все выделенные элементы'].extend(element_list)

    # separate elements into sets based on their type
    for el in element_list:
        if hasattr(el, 'LineStyle'):
            el_sets[el.LineStyle.Name].append(el)
        else:
            el_type = revit.doc.GetElement(el.GetTypeId())
            if el_type:
                el_sets[revit.query.get_name(el_type)].append(el)

    return el_sets


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    selection = revit.get_selection()
    if not selection:
        with forms.WarningBar(title="Выберите элементы"):
            selection = revit.pick_rectangle("Выберите элементы")
            if not selection:
                script.exit()

    options = process_options(selection)
    if not options:
        forms.alert("Не найдены параметры для расчета суммы.", exitscript=True)

    selected_switch = \
        forms.CommandSwitchWindow.show(sorted(options),
                                       message='Значения параметра:')

    if not selected_switch:
        script.exit()

    selected_option = options[selected_switch]
    if not selected_option:
        script.exit()

    for type_name, element_set \
            in process_sets(selection).items():
        type_name = coreutils.escape_for_html(type_name)
        output.print_md('### Итого: {}'.format(type_name))
        output_param_total(element_set, selected_option)
        output.print_md('#### Список значений:')
        output_breakdown(element_set, selected_option)
        output.insert_divider(level='##')


script_execute()
