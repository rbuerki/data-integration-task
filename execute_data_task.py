import datetime as dt
import logging
import shelve

import numpy as np
import pandas as pd

PATH_SUP = "data/supplier_car.json"
PATH_TARGET = "data/target_data.xlsx"
PATH_MAPPERS = "data/mapper_dicts"


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)


def load_json_supplier_data(path):
    """Read supplier data into a dataframe, replace "null"
    values with np.nan. Note: The JSON is in line format.
    """
    df = pd.read_json(path, lines=True)
    df.replace("null", np.nan, inplace=True)
    return df


def load_excel_target_data(path, engine="openpyxl"):
    """Read target data into dataframe, I will use it to generate
    several mapping dicts. Note: I use the openpyxl engine here.
    You can change that and use another like xlread or so.
    """
    df_target = pd.read_excel(path, engine=engine)
    return df_target


def load_prepared_mapping_dicts(path):
    """Load two prepared dictionaries. One is the quasi-mythical
    "MAIN MAPPER", defining which columns of the target data
    correspond to which columns of the supplier data. The other
    is a simple color mapper (german to english).
    (See DEV jupyter notebook for the dictionary definition.)
    """
    with shelve.open(path, "r") as shelf:
        main_mapper = shelf["main_mapper"]
        color_mapper = shelf["color_mapper"]
        return main_mapper, color_mapper


def pivot_attributes_of_supplier_data(df):
    """Pivot the `Attribute Names` and `Attribute Values` columns to
    bring the df into a tidy format (1 row per entity). Note: You
    have to make sure you have no missing values in the index columns!
    """
    # Replace the nan values temporarily with a placeholder for the pivot
    df = df.fillna("xxx").copy()
    df = df.pivot_table(
        index=[
            "ID",
            "MakeText",
            "TypeName",
            "TypeNameFull",
            "ModelText",
            "ModelTypeText",
        ],
        columns="Attribute Names",
        values="Attribute Values",
        aggfunc="max",
    ).reset_index()
    df.columns.name = None
    df = df.replace("xxx", np.nan)
    return df


def map_colors(df, color_mapper):
    """Map the colors using a mapping dictionary. If a color
    is not found in the dict, "Other" is used as default value.
    Note: We lose the `mét.` information in this process.
    """
    df = df.copy()
    df["BodyColorText_mapped"] = (
        df["BodyColorText"]
        .str.split(" ")
        .str.get(0)
        .map(color_mapper)
        .fillna("Other")
    )
    return df


def check_color_mapping(df):
    """Log the result of the color mapping. Warn if a certain
    color had to be mapped to the default value.
    """
    df = df.copy()
    defaults = df[df["BodyColorText_mapped"] == "Other"]["BodyColorText"]
    defaults = [col for col in defaults.unique()]
    if len(defaults) > 0:
        message = (
            f"CHECK! The following color(s) have been mapped to 'Other': {defaults}"
        )
    else:
        message = "All colors have been mapped to specific values."
    return message


def create_make_look_up(df):
    """Create a mapping of all make values in
    the target data frame with x.lower(): x. This
    will help to map the makes of the supplier
    data independent of the case.
    """
    make_look_up = dict(
        zip(
            [x.lower() for x in df["make"].unique().astype(str)],
            list(df["make"].unique().astype(str)),
        )
    )
    return make_look_up


def map_makes(df, make_look_up):
    """Map the makes using the look_up. If a make cannot be
    mapped we use it anyway with an appendix "_SUP". So
    we can decide what to do about it. The make info is to
    important to be defaulted.
    """
    df["MakeText_mapped"] = df["MakeText"].str.lower().map(make_look_up)
    df["MakeText_mapped"] = np.where(
        df["MakeText_mapped"].isna(),
        df["MakeText"] + "_SUP",
        df["MakeText_mapped"],
    )
    return df


def check_make_mapping(df):
    """Log the result of the make mapping. Warn if a certain
    make could not be mapped. In this case we will keep the
    original values and not map to a default value.
    """
    df = df.copy()
    not_mapped = df[df["MakeText_mapped"].str.endswith("_SUP")]["MakeText"]
    not_mapped = [col for col in not_mapped.unique()]
    if len(not_mapped) > 0:
        message = (
            f"CHECK! The following make(s) have not been mapped to pre-existing values: {not_mapped}"  # noqa: B950
        )
    else:
        message = "All colors have been mapped to specific values."
    return message


def create_columns_lists(df, df_target, main_mapper):
    """Define which columns have to be deleted, renamed (and how)
    or filled with nan values tor bring the dataframe into the
    final state.
    """
    cols_to_delete = [x for x in df.columns if x not in main_mapper.values()]
    cols_to_rename = {v: k for k, v in main_mapper.items() if v}
    cols_tbd = [k for k, v in main_mapper.items() if v is None]
    cols_target = df_target.columns.tolist()
    return cols_to_delete, cols_to_rename, cols_tbd, cols_target


def bring_df_to_target_format(
    df, cols_to_delete, cols_to_rename, cols_tbd, cols_target
):
    """Using the column lists from the last function: Bring the
    dataframe in the final format so that it can be integrated with
    the existing data.
    """
    df = df.copy()
    df = df.drop(cols_to_delete, axis=1)
    df = df.rename(columns=cols_to_rename)
    for col in cols_tbd:
        df[col] = "TBD"
    assert df.shape[1] == len(cols_target)
    df = df.reindex(cols_target, axis=1)

    df["manufacture_year"] = df["manufacture_year"].astype("int")
    df["mileage"] = df["mileage"].astype("float")
    df["manufacture_month"] = df["manufacture_month"].astype("int")
    return df


def write_to_excel(df_tidy, df_normal, df_final):
    """Write the dataframe to excel, one sheet each for the
    three processing steps. Note: Here I use the xlxswriter engine,
    and this one cannot be changed.
    """
    path = (
        f"complete_task_{dt.datetime.strftime(dt.datetime.now(), '%Y-%m-%d-%H-%M-%S')}.xlsx"  # noqa: B950
    )
    writer = pd.ExcelWriter(path, engine="xlsxwriter")
    df_tidy.to_excel(writer, sheet_name="STEP_1", index=False)
    df_normal.to_excel(writer, sheet_name="STEP_2", index=False)
    df_final.to_excel(writer, sheet_name="STEP_3", index=False)

    # Format the color of the two mapped cols in df_normal
    format_m = writer.book.add_format({"fg_color": "#f0f921"})
    writer.sheets["STEP_2"].set_column("AA:Z", None, format_m)

    # Setting col witdh to max_len of col values + 1, with a min of 15
    for sheet, df in zip(writer.sheets.values(), [df_tidy, df_normal, df_final]):
        for pos, col in enumerate(df):
            max_len_values = df[col].astype(str).map(len).max()
            len_colname = len(df[col].name)
            sheet.set_column(pos, pos, max([15, max_len_values + 1, len_colname + 1]))
    writer.save()


def main(path_sup=PATH_SUP, path_target=PATH_TARGET, path_mappers=PATH_MAPPERS):
    sup_raw = load_json_supplier_data(path_sup)
    logging.info(f"Supplier data loaded with shape: {sup_raw.shape}")
    target_data = load_excel_target_data(path_target, engine="openpyxl")
    main_mapper, color_mapper = load_prepared_mapping_dicts(path_mappers)
    sup_tidy = pivot_attributes_of_supplier_data(sup_raw)
    logging.info(f"Supplier data re-structured, new shape: {sup_tidy.shape}")
    sup_normal = map_colors(sup_tidy, color_mapper)
    message_color = check_color_mapping(sup_normal)
    logging.info(f"Colors mapped into new column.\n {message_color}")
    make_look_up = create_make_look_up(target_data)
    sup_normal = map_makes(sup_normal, make_look_up)
    message_make = check_make_mapping(sup_normal)
    logging.info(f"Makes mapped into new column.\n {message_make}")
    columns_lists = create_columns_lists(sup_normal, target_data, main_mapper)
    sup_final = bring_df_to_target_format(sup_normal, *columns_lists)
    logging.info(f"Supplier data brought to target format, new shape {sup_final.shape}")
    write_to_excel(sup_tidy, sup_normal, sup_final)
    logging.info("Success! Written to XLSX file, task complete.")


if __name__ == "__main__":
    main()
