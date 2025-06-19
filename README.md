# Monthly Invoice Summary Automator

## Overview
This tool serves to automate the generation of the monthly invoices summary file. Traditionally, individual invoices have been manually translated into a singular CSV file summarizing the overall monthly costs. This method is prone to errors, duplicated or missing entries, and typos. This tool aims to simplify the task so no or minimal manual processing is needed.

## Requirements/Notes
#### You will need:
- A Windows Machine
- A single folder containing all individual invoices
- All invoices follow the same format

## How It Works
- The tool allows you to select a folder
- Then for every invoice file within the directory, it parses the document into a single text file
- It then generates a CSV file based on the categor(y/ies) identified by the user

#### Note
This tool is currently in testing phase. We will try 2 methods:
1. direct text parsing with regex
2. area selection

