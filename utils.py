# coding: utf-8
import re


def extractMedicines(medicine_str):
    """
    Extract medicines list from medicine string.
    :param medicine_str:
    :return:
    """
    medicines = []
    if medicine_str is None or medicine_str == "":
        return medicines

    # Splite medicine string by "、"
    items = medicine_str.split("、")
    # print("item num: %d" % len(items))

    for it in items:
        result = re.findall(u'[\u4e00-\u9fff]+', it)
        if result is None or len(result) == 0:
            continue
        medicine_str = result[0]

        if "等" in medicine_str:
            medicine_str = medicine_str.replace("等", "")
        if "各" in medicine_str:
            medicine_str = medicine_str.replace("各", "")
        if "份的" in medicine_str:
            medicine_str = medicine_str.replace("份的", "")
        if "份" in medicine_str:
            medicine_str = medicine_str.replace("份", "")
        if medicine_str == "":
            continue
        medicines.append(medicine_str)
    return medicines


def extractFormOfPatent(form_str):
    """
    Extract the form list of patent
    :param form_str:
    :return:
    """
    forms = []
    if form_str is None or form_str == "":
        return forms

    # Splite form string by "、"
    items = form_str.split("、")
    # print("form item num: %d" % len(items))
    for it in items:
        if it == "":
            continue
        forms.append(it.strip())

    return forms


def extractFunctionFromPatent(function_str):
    """
    Extract functions from patent string.
    :param function_str:
    :return:
    """
    functions = []
    if function_str is None or function_str == "":
        return functions

    # Split function string by "、"
    items = function_str.split("、")
    # print("function item num: %d" % len(items))
    for it in items:
        if it == "":
            continue
        if "\n" in it:
            it = it.replace("\n", "")
        functions.append(it.strip())


    return functions


def extractMajorFunctionFromPatent(major_function_str):
    """
    Extract major functions from patent.
    :param major_function_str:
    :return:
    """
    major_functions = []
    if major_function_str is None or major_function_str == "":
        return major_functions

    # Split major functions string by "、"
    items = major_function_str.split("、")
    print("major item num: %d" % len(items))

    # process item strong
    for it in items:
        if it == "":
            continue
        if "\n" in it:
            it = it.replace("\n", "")
        major_functions.append(it.strip())

    return major_functions


def statisMedicines(medicines):
    """

    :param medicines:
    :return:
    """