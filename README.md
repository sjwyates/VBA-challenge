# DU Data Bootcamp - VBA Homework

## Overview

For this assignment, I wanted to focus on a few aspects:

- Object-oriented programming
- Data structures and algorithms

### Object-oriented programming in VBA

As someone with a little background in Java and C#, one of my first thoughts was, "Can I do this with OOP?" In particular, I wanted to write a Stock class and create a new instance for each ticker ID, and then a collection to store all the Stock objects.

As it turns out, VBA is reasonably OOP-friendly, although it took a lot of work trying to navigate its many limitations and quirks. For instance, there is a type of Module called *Class Module* - which is a good start - but it's a mixed bag:

- It doesn't support class-based inheritance, but that's ok, I wasn't going to use those anyway. (It does have interfaces, so you could potentially jury-rig an interface to act like an abstract class.)
- There are access modifiers, as well as getters and setters, so it supports abstraction/encapsulation.
- You can't pass any arguments to the constructor, so it's really just for creating/assigning a random unique ID.

Once I got my head around VBA class modules, I went ahead and built out my CStock class. (I learned that it's important to throw extra letters on things as a naming convention - ie, *theStock* - because VBA gets easily confused and starts lower-casing things for no reason.) It was fairly straightforward, just a lot of boilerplate, like getters and setters.

Next, I wanted to see if VBA had anything resembling Java's Collection class, or any of its subclasses. Turns out VBA does have a built-in Collection class, and it does some things a Java collection can do:

- It has Add and Remove methods, an Item method to retrieve an object, and a Count property.
- The Add method also allows you to pass in a key as an optional parameter, and then pass the key instead of an index to Item or Remove. So if you know the key, you can do a lookup instead of a loop to find the object.
- Unlike arrays, you can't specify a type. Everything you add to it ends up as a Variant, which carries some risk since you have no control over what ends up in there. It also means primitives get converted to variants, so a collection of integers is much larger than an array of integers.
- There's no way to look up an item in a collection to see if it exists, except to write a function that calls the Item method, then returns True or False based on whether it throws an error. Or you can use a loop, but the error handler implementation is actually a fair bit more elegant.
- Speaking of loops, you can actually iterate over collections using For Each, which was a nice surprise, and a big advantage over arrays.

There is a way to write your own custom collection classes to get around this, but it turns out it's completely bonkers, so I abandoned that idea. Ultimately I settled on using the built-in Collection class, using the optional *key* parameter to make it act like a linked list.

### Implementation

#### Step 1: Build the stock class

Once I wrapped my head around OOP in VBA, I created a class module called *CStock* to contain the data for each stock, for each year. Here's a high-level overview of how it works:

- I declared a bunch of private fields (ticker ID, year, initial value/date, final value/date, total volume)
- I created getters for each field, plus 2 additional getters that calculate and return yearly change and percent change
- The setters for ticker ID and year are straightforward
- The total volume setter increments total volume by the amount passed to it
- There are "double setters" for initial value/date and final value/date that take in both those values, then do the logic to determine whether to update those values (**NOTE: still need to implement this**)
- Constructors in VBA can't accept arguments, so I didn't have any use for it

#### Step 2: The first loop

Once I had the *CStock* class built out, I created the subroutine. If you ignore all the formatting spaghetti, you'll see it mainly consists of 2 separate loops. The first one gets all the data from the worksheet:

1. Instantiate a Collection object called *theStocks*
2. Loop through every row in the table
3. For each row, derive the *year* from the *date* using a little util function 
4. Then concatenate the ticker ID and year - eg, AA_2016 - to use as a unique ID for each stock
5. Set the *exists* variable to the return value of *ExistsInCollection* function, which takes in the ID and *theStocks*, then passes the ID to the *Item* method on *theStocks*
6. If calling *Item* throws an error, it means the stock doesn't exist, so *ExistsInCollection* returns *False*; if no error is thrown, it means the stock already exists, so *ExistsInCollection* returns *True*
7. If *exists* now equals *False*, instantiate a new *CStock* object called *theStock*, set its initial values, and finally call the *Add* method on *theStocks* and pass it *theStock*, as well the unique ID
8. If *exists* now equals *True*, call *InitialValueAndDate*, *FinalValueAndDate*, and *IncTotalVolume* and let the *CStock* object determine whether and how to update those fields

That's how the first loop works. Note that by assigning each stock a key in the collection, I was able to use a simple lookup instead of a nested loop with logic in it, which is significantly faster.

#### Step 3: The second loop

During the first loop, we didn't actually render anything to the worksheet itself. That was a deliberate choice: Rendering is expensive, so I had it do the bulk of the processing using a simple data structure in memory. Once that's done, we have a nice tidy collection (*theStocks*) to loop over.

The second loop does the rendering for each stock, plus some logic to determine greatest percent increase, etc logic for the bonus problem. And because we're looping over a collection, we can use a *For Each* loop

#### Formatting

The formatting is where most of the spaghetti is, just because there's a lot in there. The requirements for the assignment account for a fairly small fraction. It's mostly presentational concerns: Certain columns need to be wider than others, headers should be shaded, tables should have borders, things look better centered, etc.

### Demonstration:

Here's the end product. As you can see, it takes about 15 seconds to process a fairly large dataset (700k rows):

![Demonstration](images/vba-stock-demo-syates.gif)