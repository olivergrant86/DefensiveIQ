REQUIRED_COLUMNS = [

    "Down",

    "Distance",

    "Formation",

    "Personnel",

    "Play Type",

    "Gain",

]


def validate_columns(columns):

    missing = []

    for column in REQUIRED_COLUMNS:

        if column not in columns:
            missing.append(column)

    return missing