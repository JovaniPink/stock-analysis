# Stock Analysis

> Playing with VBA for automating data analysis of stock data calculations.

## Table of Contents

- [Overview](#overview)
- [Results](#results)
- [Summary](#summary)
- [Extra](#extra)
- [Todo Checklist](#todo-checklist)
- [Contributing](#contributing)
- [License](#license)

## Overview of Project

We wanted to help a friend do some analysis on some stocks that his client were interested in. We were given the stock information of two years (2017 & 2018) of 'Green Stocks.'

We first looked through the data and focused on one stock ticker to work out the logic of the script we wanted to automate repetitive tasks of:

    - Find the total volume for the current ticker in the row.
    - Find the starting price for the current ticker.
    - Find the ending price for the current ticker.
    - Find the percentage change of the starting and ending price.

We had to loop through through the 3000 plus rows of stock data to out put the ticker volume and find the total volume for the current ticker. However, it just wasn't one ticker stock that we were looking into - we had many. So investing time in developing a VBA macro script would bring us value.

## Summary

Committed in our shared repo is the refactored VBA script (green_script_VBA.vbs) we used to create:

An analysis of specific yearly stock data.

It calculates the stock price change for the year, the price percent change for the year, and the total volume traded for that year.

It will then sort through the calculated data and determine the best performer, the worst performer, and the stock with the greatest volume traded for the year.

Our VBA script will take in a year and will cycle through each sheet automatically.

### Technologies

- Excel [Excel developer documentation](https://developer.microsoft.com/en-us/excel/docs)
- VBA [Getting started with VBA in Office](https://docs.microsoft.com/en-us/office/vba/library-reference/concepts/getting-started-with-vba-in-office)

### Resources

#### 2017

![2017](resources/2017.png)

#### 2018

![2018](resources/2018.png)

## Extra

Beyond VBA ... I wanted to translate the VBA macros to the new JavaScript API:

https://github.com/JovaniPink/stock-analysis-typescript

- [Excel add-ins documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/)
- [Excel JavaScript API overview](https://docs.microsoft.com/en-us/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
- [Work with worksheets using the Excel JavaScript API](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-worksheets)
- [Work with tables using the Excel JavaScript API](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-tables)
- [Work with ranges using the Excel JavaScript API](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-ranges)

## Todo Checklist

A helpful checklist to gauge how your README is coming on what I would like to finish:

- [ ] Fill in the three major analysis sections. :)
- [ ] Create a new worksheet that pulls in the CSV files as data to be worked on.
- [ ] Use the 'green_stocks_javascript.xlsx' as the Excel JavaScript API playground.

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.

1. Fork this repository;
2. Create your branch: `git checkout -b my-new-feature`;
3. Commit your changes: `git commit -m 'Add some feature'`;
4. Push to the branch: `git push origin my-new-feature`.

**After your pull request is merged**, you can safely delete your branch.

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for more information.
