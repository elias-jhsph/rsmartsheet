# rsmartsheet

R SDK for Smartsheet API

## Description

`rsmartsheet` is an R package that provides an interface to the Smartsheet API. It allows you to interact with Smartsheet programmatically, enabling you to perform various operations such as downloading sheets, managing attachments, creating new sheets from CSV files, and more.

## Features

- Set and check Smartsheet API key
- Set a working folder in Smartsheet
- Get Smartsheet folder ID
- Create new sheets from CSV files
- Get sheet column information
- Set sheet column types to TEXT_NUMBER
- Delete sheets by name
- Get sheet data as CSV
- Download and replace sheet attachments
- List accessible sheets, workspaces, and reports
- Get cell history
- Colorize sheets based on a "HEX_COLOR" column

## Installation

You can install the `rsmartsheet` package from GitHub using the following command:

```R
devtools::install_github("elias-jhsph/rsmartsheet", upgrade="never")
```

## Usage

First, set your Smartsheet API key using the `set_smartsheet_api_key()` function:

```R
set_smartsheet_api_key("YOUR_API_KEY")
```

Then, you can use the various functions provided by the package to interact with Smartsheet. Here are a few examples:

```R
# Create a new sheet from a CSV file
csv_to_sheet("path/to/file.csv")

# Get sheet data as CSV
sheet_csv <- get_sheet_as_csv("Sheet Name")

# Delete a sheet by name
delete_sheet_by_name("Sheet Name")

# List accessible sheets
sheets <- list_sheets()
```

For more detailed usage instructions and examples, please refer to the function documentation.

## Configuration

To use the `rsmartsheet` package, you need to have a valid Smartsheet API key. You can obtain an API key from your Smartsheet account settings.

Set the API key using the `set_smartsheet_api_key()` function before using any other functions in the package.

## Contributing

Contributions to the `rsmartsheet` package are welcome! If you find any issues or have suggestions for improvements, please open an issue or submit a pull request on the GitHub repository.

## License

This package is released under the MIT License. See the LICENSE file for more details.

## Contact

For any questions or inquiries, please contact the package maintainer:

Elias Weston-Farber
Email: eweston4@jhu.edu
