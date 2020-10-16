# Helix ALM Item Linking Script

## Purpose

Helix ALM features robust record importing. However, importing link relationships among the newly imported records can be challenging. This script is designed to read an Excel sheet which contains a list of unique record identifiers and their associated Helix ALM record type, and perform automatic linking of those records.

The script will facilitate single Parent/Child links, and single Peer to Peer links. It does not group items into a single pooled link.

## Getting Started

There are five variables needed to use the script:

**BASEURL** sets the location of your Helix ALM REST API

**APIKEY** sets the value of your Helix ALM API Key and Key Secret

**PROJECTID** sets the value of your Helix ALM project id

**WORKBOOK** sets the value of the Excel workbook to use as a source for linking information

**HELIXFIELD** sets the value for the Helix ALM custom field used to determine which records should be linked

Within your Excel file, your data should be in the following format:

### Column A: Link Type
This column determines whether to link with a Peer relationship, or a Parent/Child relationship. The only two accepted values are 'Peers' and 'ParentChild', without quotes.

**Please Note:** The 'Peers' or 'ParentChild' setting in this column must match the configuration of the Link Name in column B.

### Column B: Link Name
This column determines the specific link to use when linking the items listed on this row. You may use any link you have defined in Helix ALM, but its configuration must match the link type defined in column A.

### Column C: Parent/Peer Item Type
This column determines the type of the first linked item. If it is for a parent/child relationship, this should be the parent's item type. If it is for a peer relationship, this should be the item type for the item identified in column D.

### Column D: Parent/Peer Identifier
This column contains the unique identifier used to identify the first linked item. If it is for a parent/child relationship, this is the parent's ID. If it is for a peer relationship, this is for the item type defined in column C.

The unique identifier field used by the script is configurable. The field name you create in Helix ALM to hold this identifier should be set in the "HELIXFIELD" variable in the LinkItems.py script.

## Column E: Child/Peer Item Type
This column determines the type of the second linked item. If it is for a parent/child relationship, this should be the child's item type. If it is for a peer relationship, this should be the item type for the item identified in column F.

## Column F: Child/Peer Identifier
This column contains the unique identifier used to identify the second linked item. If it is for a parent/child relationship, this is the child's ID. If it is for a peer relationship, this is for the item type defined in column E.

The child unique identifier uses the same field as defined in column D.
