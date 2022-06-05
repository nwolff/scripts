#!/usr/bin/python3

"""
When reopening a word mail merge docx file. and the backing excel file is not in the expected location,
word will ask for the location of the excel file and if you provide it then it will connect,
but it will lose the filters and sorts that were set in the original docx file.

This tool shows the filter and sort information of a mailmerge docx file. If the info is found it looks like this :
SELECT * FROM /Users/wolff_n/Desktop/clients 10 ans.xlsx WHERE ((Carte_Club_4Vallées = 'oui')) ORDER BY Nom
"""

import argparse
from zipfile import ZipFile
import xml.etree.ElementTree as ET

word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def retrieve_datasource_query(filepath):
    """
    A docx file is just a zip file containing files and directories in conventional places.
    The info we are looking for is stored in the word/settings.xml file
    in the val attribute of node "settings/mailMerge/query"
    """
    with ZipFile(filepath) as docx_as_zip:
        with docx_as_zip.open("word/settings.xml") as settings:
            root = ET.parse(settings).getroot()
            query_node = root.find("w:mailMerge/w:query", {"w": word_ns})
            if query_node is not None:
                return query_node.attrib[f"{{{word_ns}}}val"]


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Retrieves a mail merge datasource information (including filters and sorts)"
    )
    parser.add_argument(
        "file",
        metavar="FILE",
        help="The path to a .docx mail merge file",
    )
    args = parser.parse_args()
    print(retrieve_datasource_query(args.file))
