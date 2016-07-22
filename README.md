# opm-reporting
Pull data from a quip document and convert to wiki, markdown, and csv formats.


# Usage
```
usage: generate_reports.py [-h] [-v] [--quip_api_key QUIP_API_KEY]
                           [--quip_doc_id QUIP_DOC_ID]
                           [config_file]

Convert a quip document to misc report formats.

positional arguments:
  config_file           The configuration file specs.

optional arguments:
  -h, --help                        show this help message and exit
  -v, --verbose
  --quip_api_key QUIP_API_KEY       API key for quip.
  --quip_doc_id QUIP_DOC_ID         Quip document ID.
```

## Examples
### Basic usage
* Download the repo.
* Update config_settings.json with your api key and doc id.
* Run the script.

```
python generate_reports.py
```
