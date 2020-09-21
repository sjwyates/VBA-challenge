# DU Data Bootcamp - VBA Homework

## Overview

For this assignment, I wanted to focus on a few aspects:

- Object-oriented programming
- Data structures and algorithms

### Object-oriented programming in VBA

As someone with a little background in Java and C#, one of my first thoughts was, "Can I use OOP?" In particular, I wanted to write a Stock class and create a new instance for each stock, and then a collection to store all the Stock objects.

As it turns out, VBA is somewhat OOP-friendly, although it took a lot of work trying to navigate its many limitations and quirks. For instance, there is a type of Module called *Class Module* - which is a good start - but it's a mixed bag:

- It doesn't support class-based inheritance, but that's ok, I wasn't going to use those anyway (although you could potentially jury-rig an interface to act like an abstract class
- There are access modifiers, as well as getters and setters, so it supports abstraction/encapsulation
- You can't pass any arguments to the constructor, so it's really just for creating/assigning a random unique ID

Once I got my head around VBA class modules, I went ahead and built out my *CStock* class. (I learned that it's important to throw extra letters on things as a naming convention - ie, *theStock* - because VBA gets easily confused and starts lower-casing things.) It was fairly straightforward, just a lot of boilerplate, like getters and setters.

Next, I wanted to see if VBA had anything resembling Java's Collection class, or any of its subclasses. Turns out VBA does have a built-in Collection class, and it does some things a Java collection can do:

- It has Add and Remove methods, an Item method to retrieve an object, and a Count property
- The Add method also allows you to pass in a key as an optional parameter, and then pass the key instead of an index to Item or Remove, so if you know the key, you can do a lookup instead of a loop to find the object
- Unlike arrays, you can't specify a type - everything you add to it ends up as a Variant, which carries some risk since you have no control over what ends up in there
- There's no way to look up an item in a collection to see if it exists, except to write a function that calls the Item method, then returns True or False based on whether it throws an error (or you can use a loop, but the error handler implementation is actually a fair bit more elegant)
- You can iterate over collections using For Each, which was a nice surprise

There is a way to write your own custom collection classes to get around the whole Variant issue, but it turns out it's completely bonkers, so I abandoned that idea. Ultimately I settled on using the built-in Collection class, using that optional *key* parameter so I could do lookups instead of loops to find individual stocks.

### Implementation

#### Step 1: Build the CStock class

Once I wrapped my head around OOP in VBA, I created a class module called *CStock* to contain the data for each stock, for each year. Here's a high-level overview of how it works:

- I declared a bunch of private fields (ticker ID, year, initial value/date, final value/date, total volume)
- I created getters for the ticker, year, and volume fields, plus getters that calculate/return yearly change and percent change
- The setters for ticker and year are straightforward, and are only used during instantiation (I would put these in the constructor and make them read-only if I could)
- Initial and final date/value are initialized in a single sub, as well as total volume
- After that, a very similar-looking sub takes in the same 4 arguments and does some logic to see whether anything should update, and then increments total volume by the amount passed in

#### Step 2: Looping through all the data

Once I had the *CStock* class built out, I created the subroutine. The first part is where I get the data from all the worksheets:

1. Instantiate a Collection object called *theStocks*
2. Loop through each worksheet using a For Each loop
3. Copy all the data from the current worksheet into a 2-dimensional array and loop over it
4. Derive year from date, then concatenate ticker and year - eg, AA_2016 - to use as a unique ID for each stock

To see if the stock is already in the collection, call the *ExistsInCollection* function and set the *exists* variable to the return value, which is a little complicated:

- *ExistsInCollection* takes in the ID and *theStocks*
- Then it passes the ID to the *Item* method on *theStocks*
- If calling *Item* throws an error, it means the stock doesn't exist, so *ExistsInCollection* returns *False*; if no error is thrown, it means the stock already exists, so *ExistsInCollection* returns *True*

Based on the new value of *exists*, we do one of 2 things:

- If it's *False*, instantiate a new *CStock* object called *theStock*, set its initial values, and finally call the *Add* method on *theStocks* and pass it *theStock*, as well the unique ID
- If it's *True*, call its *UpdateValues* method and let the *CStock* object determine whether and how to update those fields

That's how the first part works. A few things to note:
- Copying the data into a 2-dimensional array first, then looping over the array is MUCH faster than looping over the rows in the spreadsheet
- By indexing each stock in the collection by key, I was able to use a simple lookup instead of a nested loop with logic in it, which is another significant performance boost

#### Step 3: Looping through the collection

At this point, nothing has actually been rendered to the worksheet itself. That was a deliberate choice: Rendering is expensive, so I wanted to make that the very last step. Now we have all the data ready to go, it's time to render.

But first, for the bonus problem I instantiated some *CStock* objects for each category, then looped over *theStocks* using a For Each loop, each time checking to see whether that stock beats any of the current record-holders. Then I looped over *theStocks* again to render the main table. (I initially had both in the same loop, but decided to separate out that logic for neatness.) Then I populated the greatests table.

#### Formatting

Lastly, some housekeeping. I wrote a few loops to populate all the header and category titles, using a neat little trick I found to easily create an array of strings with the *Split* method. Then the formatting:

1. Quick global reset: Align everything center and do an AutoFit to make it more readable
2. Add some table borders, as well as some shading and text wrapping for the headers
3. Add conditional formatting to the yearly change column, then delete that formatting from the header cell

### Demonstration:

And that's it! Here's the end product, with a little timer thrown in to demonstrate performance. takes a little under 9 seconds to process a fairly large dataset (~23 million rows spread across 3 worksheets), but with the screen recorder running it was more like 14 seconds:

![Demonstration](images/vba-stock-demo-syates.gif)

