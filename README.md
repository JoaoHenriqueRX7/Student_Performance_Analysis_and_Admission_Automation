# Student Performance Analysis and Admission Automation

## Introduction
 This repository hosts a comprehensive Python application designed to streamline the process of analyzing student performance and automating the generation of admission letters. Utilizing a robust ETL pipeline, the application processes student examination data, assesses eligibility based on predefined criteria, and produces personalized admission letters for successful candidates. The results are neatly compiled into a formatted Excel workbook, complete with hyperlinks for easy navigation.

## Features:
- Data Extraction and Loading: Integrates with Kaggle to download a dataset containing student performance metrics.
- Data Transformation: Implements Python's powerful pandas library to clean and transform the data, ensuring it meets the necessary standards for processing.
- Performance Assessment: Evaluates student scores against set benchmarks to determine eligibility for admission.
- Automated Letter Generation: Utilizes docxtpl to create personalized admission letters for each successful student, storing them in a designated directory.
- Excel Dashboard Creation: Leverages openpyxl to craft a styled Excel dashboard that provides a user-friendly interface to view student performance and access individual admission letters.

## How to Use:
1- Clone the repository to your local machine.
2- Ensure you have the necessary Python libraries installed (pandas, openpyxl, faker, and docxtpl).
```sh
pip install pandas openpyxl faker seaborn docxtpl
```
3- Run the penman.py script to start the automated data processing and letter generation.
4- Access the output Excel file in the specified output directory.

## project_root/

```sh
project_root/
│
├── data/                         # Directory for storing datasets and related files
│   ├── downloaded_data/          # Store downloaded data files here
│   └── processed_data/           # Store processed or transformed data here
│
├── notebooks/                    # Jupyter notebooks
│   └── ETL.ipynb                 # Your main ETL Jupyter Notebook
│
├── scripts/                      # Additional Python scripts (if any)
│
├── logs/                         # Directory for log files
│   └── etl_logs.txt              # Log file for ETL processes
│
├── requirements.txt              # Required Python libraries
│
└── README.md                     # Project README file

```

## Requirements:
- Python 3.x
- Kaggle API credentials placed in project_assets/kaggle.json

## License

The source code is available under the MIT license. See LICENSE for more information.

## Acknowledgments

This project was inspired by various resources and similar projects in the field of data science. Special thanks to all contributors and the open-source community.

© Copyright 2023 João Henrique. All rights reserved.

