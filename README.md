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

### Data structures and algorithms

I'm by no means an expert in this field, but I do know that lookups are faster than loops, hence the linked list option. I went in with a basic process map in my head.

1. Instantiate a Collection object
2. Loop through every row in the table
3. For each row, derive the *year* from the *date* using a little util function 
4. Then concatenates the ticker ID and year - eg, AA_2016 - to use as the unique ID as the key for each stock when it gets added to the collection
5. At this point, we're ready to check to see if the stock already exists, which is where that weird call Item/. If it doesn't, instantiate a new *CStock* object. If it does, do some logic to see if it needs to be updated.

### Demonstration:

Here's the end product. As you can see, it takes about 15 seconds to process a fairly large dataset (700k rows):

![Demonstration](images/vba-stock-demo-syates.gif)