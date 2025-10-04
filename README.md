# Secret Santa

Python script to generate a weighted draw for a Secret Santa and automatically send an email to everyone.

## Environment

This project runs using Python. I use `conda` to create an environment that includes all the required packages.
You can create the conda environment with the [environment file](./environment.yml) by running :

```bash
conda env create -f environment.yml
```

## Weighted draw

The algorithm aims to generate a weighted draw from an Excel file containing the name and email address of all the people to be considered, as well as the number of times each person has given a gift to all the other people. The weighted draw increases the probability of choosing someone you haven't chosen often, without prohibiting the same person from being drawn twice in a row.

### Excel file

The excel file uses mainly a tab *occurrence* to count the number of times each person has offered to other people. You can find a [template](./template.ods) to use as a start.

The table is in the following format : 

| name        | mail                 | Dipsy | Laa-Laa | Po  | Tinky Winky |
| ----------- | -------------------- | ----- | ------- | --- | ----------- |
| Dipsy       | dispy@mail.com       | 0     | 1       | 0   | 2           |
| Laa-Laa     | laalaa@mail.com      | 0     | 0       | 2   | 3           |
| Po          | po@mail.com          | 3     | 0       | 0   | 1           |
| Tinky Winky | tinky.winky@mail.com | 0     | 2       | 2   | 0           |

In this example, Laa-Laa offered a gift 3 time to Tinky Winky and 2 times to Po, and never picked Dipsy.

NB : In this table, the names on the columns and the rows must be in the same order to consider the occurrences properly.

There also is a tab *avoidance* to state if someone should never pick someone else in particular. Simply add the names in the table on the excel like this :

| Name  | To avoid |
| ----- | -------- |
| Dipsy | Po       |

### Probability weighting

The weighted draw is generated from the tab *occurrence* from the excel. It picks the names in a random order and for each person the algorithm :

- counts the maximum number of time the person picked each person, and add 1 to this number (called $maxOcc$)
- create a list of possibility by removing the names already picked and the people in the *avoidance* tab
- each name will be added $maxOcc - occ[name]$ times in the list

Given this principle, Tinky Winky offered 2 times a gift to Laa-Laa and Po, but never offered a gift to Dipsy. So for this example $maxOcc = 2 +1 = 3$ and his possibility list would be : *[Dipsy,Dipsy,Dipsy,Laa-Laa,Po]*.

One name is then chosen randomly from this list, thus increasing the probability of picking someone you didn't pick often but still allowing to pick other people.

## Email sending

The email sending part of this script uses [this tutorial](https://developers.google.com/gmail/api/quickstart/python) to send emails with a Gmail address using the Python API.

Follow this tutorial to set up a Google Cloud project and Gmail to enable automatic email sending.