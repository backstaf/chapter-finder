# Chapter-finder
Finds the nearest matching chapters (or paragraphs) to search terms of arbitrary length in Word documents

## Summary

The project takes any Word document and finds the most appropriate article/chapter/paragraphs based on search terms or phrases. The project is submitted as the Building AI course project, and uses a tf-idf algorithm and nearest neighbor to find the closest matches. 

## Background

Users can have a lot of difficulty finding the correct information when searching e.g. a large user manual. This means frustration and waste of time for the user.

## How is it used?

The project can be used in a very wide variety of applications and in different languages since it builds the vocabulary from the document(s) to be searched. Could be very helpful for large amounts of data in knowledge bases, manuals, etc. For the first iterations of the project both the input document(s) and search parameters are entered directly in the code before running.


<img src="https://github.com/backstaf/chapter-finder/blob/main/screenshot_chapter-finder.PNG" width="600">

## Data sources and AI methods
The data used for the AI finder is the input docs provided by the users themselves. The methods are nearest neighbor with Term Frequency Inverse Document Frequency (tf-idf)

## Challenges

Since a lot of interesting data is found outside the bounds of Word documents, this project is kind of specialized to that kind of environment. It would take just a bit of re-work to fit the project to other types of input. The project in the early stages also only takes into account text in paragraphs, and not in e.g. tables within word documents. 

## What next?

Some of the next steps could be pretty straightforward:
* Save the vocabulary and tf-idf data into files in order not to have to calculate it at every run, and implement a trigger for when to recalculate
* Include handling tables in word docs
* Generalize to other text file formats
* Tamper with prioritizing chapters/paragraphs which have the exact search phrase, since they probably are more likely to be what the user wants
* Creating a nicer UI
* Creating a way for the user to give the input search parameters on the fly without having to add it in the code base

## Acknowledgments

* This project was inspired through a real-world scenario of working with and providing support for a product with a 100+ page user manual which can be difficult to search
* In the code, the python libraries numpy and python-docx are used
