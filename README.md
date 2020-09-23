# DU Data Bootcamp - VBA Homework

## Overview

The purpose of this assignment was to build a VBA script to parse a moderately large set of data containing the daily opening and closing prices for over 3000 stocks over 3 years, comprising 2.3 million rows spread across 3 worksheets.

The goal is to determine the yearly change and total volume for each stock and display them in a table, then add some conditional formatting to show the winners and losers. The bonus challenge was to create a second table that displays the biggest winner and loser in terms of % change, as well as which stock had the largest total volume.

For this assignment, I wanted to focus on a few things:

- Object-oriented programming
- Data structures and algorithms
- Readability and economy of code

Before I go through the many details, here's the end product, with a little timer thrown in to demonstrate performance and a "status bar" so you can see what it's doing:

![Demonstration](images/vba-stock-demo-syates.gif)

As you can see, it takes a little under 12 seconds to process a fairly large dataset (although without the screen recorder running it's more like 9 seconds). To achieve this, I went a little deeper into VBA than I probably should have, but I really enjoyed working on it. If you're interested, here's an overly detailed explanation on how I approached the problem.

## Object-oriented programming in VBA

As someone with a little background in Java and C#, one of my first thoughts was, "Can I use OOP?"

As it turns out, VBA is somewhat OOP-friendly, although it took a lot of work trying to navigate its many limitations and quirks. For instance, there is a type of Module called *Class Module* - which is a good start - but it's a mixed bag:

- It doesn't support class-based inheritance, although I wasn't going to use that anyway
- There are access modifiers, as well as getters and setters, so it supports abstraction/encapsulation
- You can't pass any arguments to the constructor, so you have to use setters to set initial values

I also wanted to see if VBA had anything resembling Java's *Collection* class, or any of its subclasses. Turns out it has one, and it does some things Java collections can do, with some glaring exceptions:

- It has *Add* and *Remove* methods, an *Item* method to retrieve an object, and a *Count* property
- The *Add* method also allows you to pass in a key as an optional parameter, so you can use a lookup instead of a loop to find a particular object 
- There's no direct way to look up an item in a collection to see if it exists, so you have to use a workaround where you call the *Item* method and see if it throws an error
- Unlike arrays, you can't specify a type, so everything you add to it ends up as a *Variant*, which carries some risk since you have no control over what ends up in there

There is a way to write your own custom collection classes to get around the whole *Variant* issue, but it turns out it's completely bonkers, so I abandoned that idea and went with the built-in Collection class.

## Implementation

### Step 1: Build the *CStock* class

Once I wrapped my head around OOP in VBA, I created a class module called *CStock* to contain the data for each stock, for each year. Here's a high-level overview of how it works:

- I declared a bunch of private fields (ticker ID, year, initial value/date, final value/date, total volume)
- I created getters for the ticker, year, and total volume fields, plus getters that calculate/return yearly change and percent change
- Because the constructor can't take arguments, I had to use setters to initialize ticker ID, year, total volume, initial value/date, and final value/date
- The initial setters for ticker and year are straightforward, but for the rest I created a single sub that sets them all at once
- Lastly, I wrote a sub that takes in date, opening/closing values, and total volume that does some logic to see if any date/value fields need to be updated, then increments total volume

### Step 2: Looping through all the data

Once I had the *CStock* class built out, I created the subroutine to process the data. The first part is where I get the data from all the worksheets:

1. Instantiate a *Collection* object called *theStocks* and *Dim* a *Worksheet* object called *CurrentWS*
2. Loop through each worksheet using a *For Each* loop and set *CurrentWS* to it
3. Copy all the data from the current worksheet into a 2-dimensional array and loop over it with a *For Next* loop
4. Derive year from date, then concatenate ticker and year - eg, AA_2016 - to serve as a unique ID for each stock to keep multiple years for the same stock separate
5. Use that ID to find out if the stock exists in the collection yet by calling the *ExistsInCollection* util function, then set *exists* to its return value

That part's a bit tricky, so I'll break it down:

- *ExistsInCollection* takes in a *key* and a *Collection* object - in this case, the ID and *theStocks* - then passes the key to the *Item* method on the collection
- If calling *Item* throws an error, the stock doesn't exist yet, so the function returns *False* from the error handler
- If no error is thrown, the stock already exists, so the function returns *True* and exits

Based on the new value of *exists*, we do one of 2 things:

- If it's *False*, instantiate a new *CStock* object called *theStock*, set its initial values, and finally call the *Add* method on *theStocks* and pass it *theStock*, as well the unique ID
- If it's *True*, call its *UpdateValues* method and let the *CStock* object determine whether and how to update those fields

That's how the first part works. A few things to note:

- Copying the data by value into a 2-dimensional array first, then looping over the array is MUCH faster than looping over *Range* objects in the spreadsheet
- By indexing each *CStock* object in *theStocks* by key, I was able to use a simple lookup instead of a loop, which shaves off a bit of time as well

### Step 3: Analyzing and rendering the collection

At this point, nothing has actually been rendered to the worksheet itself. That was a deliberate choice: Rendering is expensive (similar to how looping over the sheet is more expensive than looping over an array) so I wanted to make that the very last step.

But first, for the bonus problem I instantiated some *CStock* objects for each category, then looped over *theStocks* using a For Each loop, each time checking to see whether that stock beats any of the current record-holders.

Next, I looped over *theStocks* again to populate values in the main table. (I initially had both in the same loop, but decided to separate out that logic for neatness.) To make the data a little more organized, I sorted the main table by ticker, then year, so that you could see what happened to a particular stock on consecutive years.

Then I populated the record-holders table. To make things a little less hard-coded, I added all the record-holders to a new collection so that I could loop over it, rather than repeat the same "this goes here, that goes there" code 9 times.

And lastly, I added the header and category titles, using a neat little workaround to easily create an array of strings with the *Split* method as a sort of pseudo-array literal. (VBA supposedly has array literals, but I struggled to get those to work.)

### Step 4: Formatting

Now the most important part: Make it look nice, starting with the basic stuff. I aligned everything center to make it more readable, then added table borders, as well as some shading in the headers. I also opted to AutoFit the table body contents to make sure every value looks like it was supposed to (no one likes to see a bunch of ####### and scientific notation randomly mixed in), and added WrapText for the headers (no one likes columns that are way wider than they need to be.)

Next, the conditional formatting. Initially I just used a simple boolean (more than zero = green, less than zero = red) for the yearly change, but decided it was pretty glaring and not very meaningful, so I went with a 3-color scale for the percent change instead. It was a bit more difficult, but it made a huge improvement.

### Summary

I probably went a little too crazy with this assignment, but I'm really happy with the outcome. Besides making it fast and look nice, I also sought to make the code fairly clean and organized, and implementing OOP was a big part of that. VBA throws you a lot of curveballs, but that only makes it more satisfying when you get it to work.

Feel free to reach our if you have any questions, comments, suggestions, etc.
