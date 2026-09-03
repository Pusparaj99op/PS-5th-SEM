# -*- coding: utf-8 -*-
import os
from docx import Document
from docx.shared import Pt

ROOT = os.path.join(os.path.dirname(__file__), "..")

PRACTICALS = [
{
"num": 1, "title": "Study and Implementation of Advanced Data Structures in Python",
"aim": "To study and implement advanced data structures such as Stack, Queue, and Linked List in Python.",
"theory": "A Stack is a linear data structure that follows the LIFO (Last In First Out) principle, where insertion and deletion happen at one end called the top. The two primary operations are push (insert) and pop (remove), both performed in O(1) time; stacks are used in function call management, expression evaluation, undo mechanisms, and backtracking algorithms. A Queue follows the FIFO (First In First Out) principle, where insertion (enqueue) happens at the rear and deletion (dequeue) at the front; unlike a plain Python list, collections.deque is implemented as a doubly linked list internally, giving O(1) time complexity for operations at both ends instead of O(n) for a list. A Linked List is a linear collection of nodes where each node stores data and a reference (pointer) to the next node, allowing dynamic memory allocation and efficient insertion/deletion (O(1) once the position is known) unlike arrays, which require contiguous memory and O(n) shifting. Variants include singly linked lists (one-directional traversal, used here), doubly linked lists (traversal in both directions), and circular linked lists (last node points back to the first). Together, these three structures form the foundation for more advanced data structures such as trees, graphs, and hash tables.",
"algorithm": [
"Create an empty list to act as a stack.",
"Push elements onto the stack using append() and pop the last element using pop().",
"Create a deque to act as a queue.",
"Enqueue elements using append() and dequeue using popleft().",
"Create a linked list using nested dictionaries with 'data' and 'next' keys.",
"Traverse the linked list from head to tail, collecting values until next is None.",
"Display the results of all three data structures.",
],
"code": """stack = []
stack.append(10)
stack.append(20)
stack.append(30)
print("Stack:", stack)
print("Popped:", stack.pop())

from collections import deque
queue = deque()
queue.append("A")
queue.append("B")
queue.append("C")
print("Queue:", list(queue))
print("Dequeued:", queue.popleft())

node = {"data": 1, "next": {"data": 2, "next": {"data": 3, "next": None}}}
values = []
cur = node
while cur:
    values.append(cur["data"])
    cur = cur["next"]
print("Linked List:", values)""",
"output": """Stack: [10, 20, 30]
Popped: 30
Queue: ['A', 'B', 'C']
Dequeued: A
Linked List: [1, 2, 3]""",
"result": "The stack, queue, and linked list data structures were successfully implemented and their basic operations (push/pop, enqueue/dequeue, traversal) were verified.",
},
{
"num": 2, "title": "Object-Oriented Programming Concepts",
"aim": "To study and implement Object-Oriented Programming concepts such as classes, objects, inheritance, and method overriding in Python.",
"theory": "Object-Oriented Programming (OOP) is a programming paradigm based on the concept of objects, which bundle data (attributes) and behavior (methods) together, in contrast to procedural programming which separates data and functions. OOP rests on four main pillars: Encapsulation, which hides internal state and requires interaction through well-defined methods; Abstraction, which exposes only essential features while hiding implementation detail; Inheritance, which allows a class (child/subclass) to acquire properties and methods of another class (parent/superclass), promoting code reuse; and Polymorphism, which allows objects of different classes to be treated through a common interface. A class is a blueprint or template for creating objects, while an object is a concrete instance of a class with its own state. The __init__ method is a constructor, automatically invoked when an object is created, used to initialize instance attributes. Method overriding is a form of runtime polymorphism where a subclass provides a specific implementation of a method already defined in its parent class; when the method is called on a subclass object, Python's dynamic dispatch ensures the overridden version executes, which is demonstrated when Dog and Cat objects both respond differently to the same speak() call.",
"algorithm": [
"Define a base class Animal with an __init__ method and a speak() method.",
"Define subclasses Dog and Cat that inherit from Animal.",
"Override the speak() method in each subclass to return a specific sound.",
"Create a list of Animal objects containing Dog and Cat instances.",
"Loop through the list and call speak() on each object to demonstrate polymorphism.",
],
"code": """class Animal:
    def __init__(self, name):
        self.name = name

    def speak(self):
        return f"{self.name} makes a sound"

class Dog(Animal):
    def speak(self):
        return f"{self.name} barks"

class Cat(Animal):
    def speak(self):
        return f"{self.name} meows"

animals = [Dog("Rex"), Cat("Whiskers")]
for a in animals:
    print(a.speak())""",
"output": """Rex barks
Whiskers meows""",
"result": "OOP concepts including class definition, inheritance, and method overriding were successfully implemented and verified using Python.",
},
{
"num": 3, "title": "File Handling and OS Module Operations",
"aim": "To study and implement file handling and OS module operations such as creating directories, reading/writing files, renaming, and deleting files in Python.",
"theory": "File handling in Python allows programs to create, read, write, append to, and manipulate files stored on persistent storage using the built-in open() function, which supports modes such as 'w' (write, truncates existing content), 'r' (read), 'a' (append), and 'x' (exclusive creation), along with binary variants like 'wb'/'rb'. Using a 'with' statement (context manager) ensures the file is automatically and safely closed even if an exception occurs, which is why it is preferred over manual open()/close() calls. The os module provides a cross-platform interface to interact with the underlying operating system: os.makedirs() creates a directory (and any missing intermediate directories, with exist_ok=True avoiding an error if it already exists), os.listdir() lists the names of entries in a directory, os.rename() renames or moves a file, os.remove() deletes a file, and os.rmdir() deletes an empty directory. Related modules such as os.path (for path manipulation and existence checks) and shutil (for recursive copy/move/delete of non-empty directories) are commonly used alongside the os module for complete file-system automation.",
"algorithm": [
"Import the os module.",
"Create a directory using os.makedirs().",
"Open a file in write mode and write text into it, then close it.",
"Open the same file in read mode and display its content.",
"List all files in the directory using os.listdir().",
"Rename the file using os.rename() and list the directory again.",
"Delete the file using os.remove() and delete the directory using os.rmdir().",
],
"code": """import os
os.makedirs("demo_folder", exist_ok=True)

with open("demo_folder/sample.txt", "w") as f:
    f.write("Hello, File Handling!\\n")
    f.write("This is line two.")

with open("demo_folder/sample.txt", "r") as f:
    print("File Content:")
    print(f.read())

print("Files in folder:", os.listdir("demo_folder"))

os.rename("demo_folder/sample.txt", "demo_folder/renamed.txt")
print("After rename:", os.listdir("demo_folder"))

os.remove("demo_folder/renamed.txt")
os.rmdir("demo_folder")
print("Cleanup done")""",
"output": """File Content:
Hello, File Handling!
This is line two.
Files in folder: ['sample.txt']
After rename: ['renamed.txt']
Cleanup done""",
"result": "File handling and OS module operations (create, read, rename, delete) were successfully performed and verified.",
},
{
"num": 4, "title": "Exception Handling and Fetching Data from a REST API",
"aim": "To study and implement exception handling in Python while fetching data from a REST API.",
"theory": "Exception handling in Python is done using try, except, else, and finally blocks: code that might raise a runtime error is placed in the try block, the except block catches and handles specific (or general) exception types so the program does not crash abruptly, an optional else block runs only if no exception occurred, and the finally block always executes regardless of whether an exception was raised, making it ideal for cleanup actions like closing files or connections. Catching specific exception classes (e.g., URLError, TimeoutError) instead of a bare except is considered good practice as it avoids silently hiding unrelated bugs. A REST API (Representational State Transfer) is an architectural style that allows applications to communicate over HTTP using standard methods such as GET, POST, PUT, and DELETE, typically exchanging data in JSON (JavaScript Object Notation) format because it is lightweight and language-independent. The urllib.request module (a built-in alternative to third-party libraries like requests) can be used to send HTTP requests and read server responses, while the json module provides json.loads() to parse a JSON string/bytes into native Python objects (dicts/lists) and json.dumps() to convert Python objects back into a formatted JSON string.",
"algorithm": [
"Import the json and urllib.request modules.",
"Define the API URL to fetch data from.",
"Use a try block to open the URL and read the response with a timeout.",
"Parse the response using json.loads() and print it in readable format using json.dumps().",
"Use an except block to catch and print any errors during the request.",
"Use a finally block to indicate that the request has completed.",
],
"code": """import json
import urllib.request

try:
    url = "https://jsonplaceholder.typicode.com/todos/1"
    with urllib.request.urlopen(url, timeout=5) as response:
        data = json.loads(response.read())
        print("API Response:")
        print(json.dumps(data, indent=2))
except Exception as e:
    print("Error fetching data:", e)
finally:
    print("Request completed")""",
"output": """API Response:
{
  "userId": 1,
  "id": 1,
  "title": "delectus aut autem",
  "completed": false
}
Request completed""",
"result": "Exception handling was successfully used to safely fetch and parse JSON data from a REST API.",
},
{
"num": 5, "title": "Python Libraries and Basic Web Scraping",
"aim": "To study and implement basic web scraping using Python libraries requests and BeautifulSoup.",
"theory": "Web scraping is the process of programmatically extracting information from websites by fetching their raw HTML and parsing out the required data, commonly used for data collection, price monitoring, and research when no structured API is available; it should always be done in compliance with a site's terms of service and robots.txt. The requests library is a popular third-party HTTP client used to send GET/POST requests and retrieve a page's HTML content along with metadata like status codes and headers, and is generally preferred over the lower-level urllib for its simpler, more readable API. BeautifulSoup, from the bs4 package, is an HTML/XML parsing library that builds a navigable parse tree from raw markup and provides convenient methods to search and traverse it, such as tag access via dot notation (soup.title, soup.h1), find() and find_all() to locate elements by tag name, class, or attribute, and .text/.string to extract the readable text content of an element, abstracting away the need to manually write regular expressions against HTML.",
"algorithm": [
"Import the requests library and BeautifulSoup from bs4.",
"Send a GET request to the target URL to fetch the HTML page.",
"Parse the HTML content using BeautifulSoup with the 'html.parser'.",
"Extract and print the page title using soup.title.string.",
"Extract and print the first heading using soup.h1.text.",
"Find all paragraph tags using soup.find_all('p') and display their count and content.",
],
"code": """import requests
from bs4 import BeautifulSoup

url = "https://example.com"
response = requests.get(url, timeout=5)
soup = BeautifulSoup(response.text, "html.parser")

print("Page Title:", soup.title.string)
print("First Heading:", soup.h1.text)
paragraphs = soup.find_all("p")
print("Number of paragraphs:", len(paragraphs))
print("First paragraph:", paragraphs[0].text)""",
"output": """Page Title: Example Domain
First Heading: Example Domain
Number of paragraphs: 2
First paragraph: This domain is for use in documentation examples without needing permission. Avoid use in...""",
"result": "Basic web scraping was successfully implemented, extracting the title, heading, and paragraph content from a web page.",
},
{
"num": 6, "title": "Arrays, Strings, Linked Lists, Stacks and Queues",
"aim": "To study and implement basic operations on arrays, strings, linked lists, stacks, and queues in Python.",
"theory": "Arrays (implemented as lists in Python) store elements in contiguous, indexable memory locations, giving O(1) random access by index; appending to the end is amortized O(1), while inserting/deleting at arbitrary positions is O(n) due to element shifting. Strings in Python are immutable sequences of characters, meaning any operation that appears to modify a string (such as reversal) actually creates and returns a new string object; slicing with a step of -1 (s[::-1]) is a concise idiom for reversal because it walks the sequence from the last index to the first. A Stack enforces LIFO (Last In First Out) order, supporting push (insert at top) and pop (remove from top) in O(1) time, useful for reversing sequences, expression parsing, and recursion simulation. A Queue enforces FIFO (First In First Out) order, supporting enqueue (insert at rear) and dequeue (remove from front); collections.deque is preferred over a plain list for queues because it is backed by a doubly linked structure, giving O(1) complexity for operations at both ends versus O(n) for list.pop(0).",
"algorithm": [
"Create a list (array) and append a new element to it.",
"Reverse a given string using slicing (s[::-1]).",
"Create a stack using a list, extend it with elements, and pop the last element.",
"Create a queue using deque, initialize it with elements, and dequeue from the front.",
"Print the results after each operation.",
],
"code": """arr = [5, 2, 8, 1]
arr.append(9)
print("Array:", arr)

s = "Hello World"
print("Reversed String:", s[::-1])

stack = []
stack.extend([1, 2, 3])
stack.pop()
print("Stack after pop:", stack)

from collections import deque
queue = deque([1, 2, 3])
queue.popleft()
print("Queue after dequeue:", list(queue))""",
"output": """Array: [5, 2, 8, 1, 9]
Reversed String: dlroW olleH
Stack after pop: [1, 2]
Queue after dequeue: [2, 3]""",
"result": "Basic operations on arrays, strings, stacks, and queues were successfully implemented and verified.",
},
{
"num": 7, "title": "Tree and Graph Data Structures with DFS and BFS",
"aim": "To study and implement graph traversal techniques - Breadth-First Search (BFS) and Depth-First Search (DFS).",
"theory": "A Tree is a hierarchical, acyclic connected graph with a single root node, where every node (except the root) has exactly one parent, commonly used to represent hierarchical relationships such as file systems and organization charts. A Graph is a more general data structure consisting of nodes (vertices) connected by edges, which may be directed or undirected, weighted or unweighted, and can contain cycles; it is often represented using an adjacency list (a dictionary mapping each node to its neighbors, as used here) or an adjacency matrix. Breadth-First Search (BFS) explores a graph level by level using a queue (FIFO), visiting all neighbors of the current node before moving to the next level, which makes it well suited for finding the shortest path in an unweighted graph. Depth-First Search (DFS) explores as far as possible along each branch before backtracking, implemented either recursively (using the call stack) or iteratively with an explicit stack, and is useful for tasks like cycle detection and topological sorting. Both algorithms maintain a 'visited' set to avoid revisiting nodes and getting stuck in infinite loops when cycles are present, and both run in O(V + E) time, where V is the number of vertices and E is the number of edges.",
"algorithm": [
"Represent the graph as a dictionary of adjacency lists.",
"For BFS: initialize a visited set and a queue with the start node; repeatedly dequeue a node, record it, and enqueue its unvisited neighbors.",
"For DFS: recursively visit the start node, mark it visited, record it, and recurse into its unvisited neighbors.",
"Call bfs() and dfs() from the same starting node and print the traversal order for each.",
],
"code": """graph = {
    'A': ['B', 'C'],
    'B': ['D'],
    'C': ['D'],
    'D': []
}

def bfs(start):
    visited, queue, order = {start}, [start], []
    while queue:
        node = queue.pop(0)
        order.append(node)
        for n in graph[node]:
            if n not in visited:
                visited.add(n)
                queue.append(n)
    return order

def dfs(node, visited=None, order=None):
    if visited is None:
        visited, order = set(), []
    visited.add(node)
    order.append(node)
    for n in graph[node]:
        if n not in visited:
            dfs(n, visited, order)
    return order

print("BFS Traversal:", bfs('A'))
print("DFS Traversal:", dfs('A'))""",
"output": """BFS Traversal: ['A', 'B', 'C', 'D']
DFS Traversal: ['A', 'B', 'D', 'C']""",
"result": "BFS and DFS graph traversal algorithms were successfully implemented and their traversal orders were verified.",
},
{
"num": 8, "title": "Sorting and Searching Algorithms",
"aim": "To study and implement Bubble Sort and Binary Search algorithms in Python.",
"theory": "Sorting algorithms arrange elements of a list into a defined order (typically ascending or descending) and are foundational to efficient searching and data processing. Bubble Sort is a simple comparison-based sorting algorithm that repeatedly steps through the list, compares each pair of adjacent elements, and swaps them if they are in the wrong order; after each full pass, the largest unsorted element 'bubbles up' to its correct position, giving a time complexity of O(n^2) in the average and worst case and O(n) in the best case (already sorted) with an early-exit optimization, and O(1) space complexity since sorting happens in place. Searching algorithms locate a target value within a collection. Linear Search checks elements sequentially (O(n)), while Binary Search is a much more efficient divide-and-conquer algorithm for finding an item in a sorted array: it repeatedly compares the target with the middle element and discards the half of the array that cannot contain the target, narrowing the search range until the element is found or the range becomes empty, achieving a time complexity of O(log n). Binary Search requires the input array to be sorted beforehand, which is why it is applied here after bubble_sort() has arranged the data.",
"algorithm": [
"Bubble Sort: for each pass, compare adjacent elements and swap them if the left is greater than the right; repeat until no swaps are needed.",
"Binary Search: set low and high pointers; compute the mid index; compare the middle element with the target; narrow the search range accordingly until found or the range is empty.",
"Sort the given array using bubble_sort().",
"Search for a target value in the sorted array using binary_search() and print its index.",
],
"code": """def bubble_sort(arr):
    n = len(arr)
    for i in range(n):
        for j in range(n - i - 1):
            if arr[j] > arr[j + 1]:
                arr[j], arr[j + 1] = arr[j + 1], arr[j]
    return arr

def binary_search(arr, target):
    low, high = 0, len(arr) - 1
    while low <= high:
        mid = (low + high) // 2
        if arr[mid] == target:
            return mid
        elif arr[mid] < target:
            low = mid + 1
        else:
            high = mid - 1
    return -1

data = [64, 25, 12, 22, 11]
sorted_data = bubble_sort(data.copy())
print("Sorted Array:", sorted_data)
print("Index of 22:", binary_search(sorted_data, 22))""",
"output": """Sorted Array: [11, 12, 22, 25, 64]
Index of 22: 2""",
"result": "Bubble Sort and Binary Search were successfully implemented, sorting the array and locating the target element correctly.",
},
{
"num": 9, "title": "Time and Space Complexity Analysis",
"aim": "To study and analyze the time and space complexity of Linear Search and Binary Search algorithms.",
"theory": "Time complexity measures how the running time of an algorithm grows as a function of input size (n), typically expressed using Big-O notation, which describes the worst-case upper bound on growth while ignoring constant factors and lower-order terms; common classes in increasing order of growth include O(1), O(log n), O(n), O(n log n), and O(n^2). Space complexity measures the additional (auxiliary) memory an algorithm requires beyond the input itself, which is important when working with large datasets or memory-constrained systems. Linear Search checks each element of the array sequentially until it finds a match or reaches the end, giving a worst-case time complexity of O(n) since it may need to examine every element. Binary Search works only on sorted data and repeatedly halves the search range by comparing the target with the middle element, giving a much better time complexity of O(log n), which scales far more efficiently as the dataset grows large - this is empirically demonstrated in this practical by timing both algorithms on a dataset of 100,000 elements. Both algorithms use O(1) additional space since they operate in-place using only a fixed number of variables regardless of input size, rather than allocating new data structures proportional to n.",
"algorithm": [
"Define linear_search() which iterates through the array checking each element against the target.",
"Define binary_search() which divides the sorted array range in half repeatedly.",
"Generate a large list of data and choose a target value.",
"Measure execution time of linear_search() using time.perf_counter().",
"Measure execution time of binary_search() similarly.",
"Print both execution times along with their theoretical time and space complexities.",
],
"code": """import time

def linear_search(arr, target):
    for i in range(len(arr)):
        if arr[i] == target:
            return i
    return -1

def binary_search(arr, target):
    low, high = 0, len(arr) - 1
    while low <= high:
        mid = (low + high) // 2
        if arr[mid] == target:
            return mid
        elif arr[mid] < target:
            low = mid + 1
        else:
            high = mid - 1
    return -1

data = list(range(100000))
target = 99999

start = time.perf_counter()
linear_search(data, target)
t1 = time.perf_counter() - start

start = time.perf_counter()
binary_search(data, target)
t2 = time.perf_counter() - start

print(f"Linear Search Time: {t1:.6f} s  (Complexity: O(n))")
print(f"Binary Search Time: {t2:.6f} s  (Complexity: O(log n))")
print("Space Complexity: O(1) for both (in-place)")""",
"output": """Linear Search Time: 0.002474 s  (Complexity: O(n))
Binary Search Time: 0.000005 s  (Complexity: O(log n))
Space Complexity: O(1) for both (in-place)""",
"result": "The time and space complexity of Linear Search and Binary Search were analyzed and empirically verified, confirming Binary Search is significantly faster on large sorted datasets.",
},
{
"num": 10, "title": "Recursion and Backtracking",
"aim": "To study and implement recursion using a Factorial program and backtracking using the N-Queens problem in Python.",
"theory": "Recursion is a programming technique where a function calls itself, directly or indirectly, to solve smaller instances of the same problem, until it reaches a base case that can be solved directly without further recursion; the factorial function demonstrates this, with factorial(0) = 1 as the base case and factorial(n) = n * factorial(n-1) as the recursive case, with each call pushing a new frame onto the call stack until the base case unwinds the recursion. Backtracking is an algorithmic technique for solving constraint-satisfaction problems incrementally, by building a solution one step at a time and abandoning ('backtracking' from) a partial solution as soon as it is determined that it cannot possibly lead to a valid complete solution, which prunes the search space far more efficiently than checking every possible combination (brute force). The N-Queens problem is a classic backtracking example: it requires placing N queens on an N x N chessboard such that no two queens attack each other (i.e., no two share the same row, column, or diagonal); the algorithm places a queen in a row, checks safety against all previously placed queens using is_safe(), recurses to the next row if safe, and backtracks (removes the queen and tries the next column) if no column in the current row is safe.",
"algorithm": [
"Factorial: if n is 0 return 1, otherwise return n multiplied by factorial(n-1).",
"N-Queens: maintain a board array storing the column position of the queen in each row.",
"Define is_safe() to check if placing a queen at a given row/column conflicts with previously placed queens.",
"Define backtrack() which tries placing a queen in each column of the current row; if safe, recurse to the next row, otherwise backtrack.",
"When all rows are filled, save the solution.",
"Print the factorial result and the number/first solution of the N-Queens problem.",
],
"code": """def factorial(n):
    if n == 0:
        return 1
    return n * factorial(n - 1)

def solve_nqueens(n):
    solutions = []
    board = [-1] * n

    def is_safe(row, col):
        for r in range(row):
            if board[r] == col or abs(board[r] - col) == abs(r - row):
                return False
        return True

    def backtrack(row):
        if row == n:
            solutions.append(board.copy())
            return
        for col in range(n):
            if is_safe(row, col):
                board[row] = col
                backtrack(row + 1)
                board[row] = -1

    backtrack(0)
    return solutions

print("Factorial of 5:", factorial(5))

solutions = solve_nqueens(4)
print(f"Number of solutions for 4-Queens: {len(solutions)}")
print("First solution (col positions per row):", solutions[0])""",
"output": """Factorial of 5: 120
Number of solutions for 4-Queens: 2
First solution (col positions per row): [1, 3, 0, 2]""",
"result": "Recursion and backtracking were successfully implemented and demonstrated using the Factorial and N-Queens programs.",
},
{
"num": 11, "title": "Responsive Web Pages using HTML, CSS and JavaScript",
"aim": "To design and implement a responsive web page using HTML, CSS, and JavaScript.",
"theory": "HTML (HyperText Markup Language) provides the semantic structure and content of a web page using nested elements (tags) such as div, h2, p, and button. CSS (Cascading Style Sheets) controls the presentation of that structure, including layout, colors, spacing, fonts, and visual effects like box-shadow and border-radius, and can be applied inline, via an internal <style> block, or via an external stylesheet. JavaScript adds interactivity and dynamic behavior to an otherwise static page by manipulating the Document Object Model (DOM) - the browser's in-memory tree representation of the HTML - at runtime, for example by changing an element's innerText or innerHTML in response to user actions. Responsive Web Design ensures a page's layout automatically adjusts to fit different screen sizes and devices (desktop, tablet, mobile) without requiring separate versions of the site, commonly achieved using relative units (%, em, rem) instead of fixed pixels, flexible containers (Flexbox/Grid), media queries, and constraints like max-width so content neither overflows small screens nor stretches awkwardly on large ones. JavaScript event handlers, such as the onclick attribute or addEventListener(), allow the page to respond dynamically to user actions like clicks, input, or form submission without requiring a full page reload.",
"algorithm": [
"Create the HTML structure with a container div styled as a 'card'.",
"Use CSS to style the body and card with background color, padding, border-radius, and box-shadow, and set a max-width on the card for responsiveness.",
"Add a button element with an onclick JavaScript handler.",
"On button click, use JavaScript to update the innerText of a paragraph element to show a message.",
"Open the page in a browser and verify the layout adjusts to screen width and the button click updates the message.",
],
"code": """<!DOCTYPE html>
<html>
<head>
<style>
  body { font-family: Arial; background: #f0f2f5; margin: 0; }
  .card { max-width: 400px; margin: 40px auto; padding: 20px;
          background: white; border-radius: 10px; box-shadow: 0 2px 8px #ccc; text-align: center; }
  button { padding: 8px 16px; background: #4CAF50; color: white; border: none; border-radius: 5px; }
</style>
</head>
<body>
  <div class="card">
    <h2>Responsive Card</h2>
    <p>This page resizes with the screen.</p>
    <button onclick="document.getElementById('msg').innerText='Button Clicked!'">Click Me</button>
    <p id="msg"></p>
  </div>
</body>
</html>""",
"output": """Browser renders a centered white card on a light grey
background with heading "Responsive Card", a paragraph,
and a green "Click Me" button.

On clicking the button, the page updates and shows:
  Button Clicked!

The card layout adjusts automatically to fit any
screen width (responsive design).""",
"result": "A responsive web page was successfully designed and implemented using HTML, CSS, and JavaScript, with verified interactivity and layout adjustment.",
},
{
"num": 12, "title": "React Basics and Component Architecture",
"aim": "To study and implement basic React concepts such as components, props, and state using a Counter application.",
"theory": "React is a declarative, component-based JavaScript library (developed by Meta) for building user interfaces out of small, reusable, self-contained pieces called components, instead of manipulating the DOM directly and imperatively as in vanilla JavaScript. A functional component is simply a JavaScript function that accepts props (inputs) and returns JSX - a syntax extension that lets HTML-like markup be written directly inside JavaScript, which is compiled (transpiled) into regular JavaScript function calls (React.createElement) by tools like Babel. Props (short for properties) are read-only data passed from a parent component to a child component, enabling component reuse with different inputs. The useState hook is one of React's built-in Hooks, which let functional components hold local, mutable state without needing class components; calling its setter function (e.g., setCount) does not mutate the variable in place but instead schedules a re-render of the component with the updated value, and React efficiently updates only the changed parts of the real DOM using its Virtual DOM diffing algorithm. Component composition, as seen with the Counter component nested inside the App component, allows complex UIs to be built by combining small, independently testable pieces, one of React's core architectural strengths.",
"algorithm": [
"Import React and the useState hook.",
"Define a Counter component that initializes a count state variable to 0 using useState.",
"Render the current count and a button that calls setCount to increment the count on click.",
"Define an App component that renders a heading and includes the Counter component.",
"Export App as the default component and render it in the browser.",
"Verify that each click on the Increment button updates the displayed count.",
],
"code": """import React, { useState } from "react";

function Counter() {
  const [count, setCount] = useState(0);
  return (
    <div>
      <h2>Count: {count}</h2>
      <button onClick={() => setCount(count + 1)}>Increment</button>
    </div>
  );
}

function App() {
  return (
    <div>
      <h1>React Basics Demo</h1>
      <Counter />
    </div>
  );
}

export default App;""",
"output": """Browser renders:
  React Basics Demo
  Count: 0
  [ Increment ]

Each click on "Increment" re-renders the Counter
component and increases the displayed count by 1:
  Count: 1
  Count: 2
  Count: 3 ...""",
"result": "Basic React concepts - components, props, and state management using useState - were successfully implemented and verified in the Counter application.",
},
{
"num": 13, "title": "Web Application using Flask Framework",
"aim": "To study and implement a simple web application using the Flask framework in Python.",
"theory": "Flask is a lightweight, minimalistic Python 'micro' web framework built on top of the Werkzeug WSGI toolkit and Jinja2 templating engine, used to build web applications and RESTful APIs with minimal boilerplate compared to full-stack frameworks like Django. A Flask application is represented by an instance of the Flask class, typically created with app = Flask(__name__), where __name__ helps Flask locate resources relative to the module. Route decorators (@app.route(path)) bind a URL path (and optionally HTTP methods like GET/POST) to a Python function known as a view function; when a client requests that path, Flask invokes the corresponding view function and sends whatever it returns (a string, JSON, or rendered HTML template) back as the HTTP response body. Multiple routes allow a single application to serve different pages or API endpoints, as demonstrated by the '/' and '/about' routes. The app.run() method starts Flask's built-in development server (not intended for production use) to listen for and serve incoming HTTP requests on a local host and port, with debug=True enabling auto-reload on code changes and detailed in-browser error tracebacks during development.",
"algorithm": [
"Import Flask and create a Flask application instance.",
"Define a route '/' mapped to a home() function that returns a welcome message.",
"Define a route '/about' mapped to an about() function that returns an about message.",
"Run the Flask application using app.run() when the script is executed directly.",
"Send GET requests to '/' and '/about' and verify the responses returned by the server.",
],
"code": """from flask import Flask

app = Flask(__name__)

@app.route("/")
def home():
    return "Welcome to my Flask Web App!"

@app.route("/about")
def about():
    return "This is the About page."

if __name__ == "__main__":
    app.run(debug=True)""",
"output": """Running: python p13.py
 * Running on http://127.0.0.1:5000 (Press CTRL+C to quit)

Tested using Flask's test client:
GET / -> 200  "Welcome to my Flask Web App!"
GET /about -> 200  "This is the About page." """,
"result": "A simple Flask web application with multiple routes was successfully created and its endpoints were verified to return the correct responses.",
},
{
"num": 14, "title": "Database Integration with Backend Application",
"aim": "To study and implement database integration with a Python backend application using SQLite.",
"theory": "SQLite is a lightweight, serverless, self-contained, file-based (or in-memory) relational database engine that requires no separate database server process or configuration, making it ideal for embedded applications, prototypes, and testing; it is bundled with Python's standard library via the sqlite3 module, so no external installation is needed. It allows applications to create tables with defined schemas (column names and data types), insert, update, delete, and query records using standard SQL (Structured Query Language) statements. A Connection object, created via sqlite3.connect(), represents an open session with the database (which can be a file path or the special value ':memory:' for a temporary, non-persistent in-memory database, as used in this practical), and is responsible for committing or rolling back transactions via conn.commit(). A Cursor object, obtained from the connection, is used to execute SQL commands (cur.execute()) and iterate over or fetch the results of queries (fetchone(), fetchall()), acting as a control structure over the result set. Backend database integration like this forms the persistence layer of most real-world applications, storing data that must survive beyond a single program execution.",
"algorithm": [
"Import the sqlite3 module and create an in-memory database connection.",
"Create a cursor object from the connection.",
"Execute a CREATE TABLE statement to define the students table schema.",
"Execute INSERT statements to add student records and commit the transaction.",
"Execute a SELECT statement to retrieve all records and print them.",
"Close the database connection.",
],
"code": """import sqlite3

conn = sqlite3.connect(":memory:")
cur = conn.cursor()

cur.execute(\"\"\"
    CREATE TABLE students (id INTEGER PRIMARY KEY, name TEXT, marks INTEGER)
\"\"\")

cur.execute("INSERT INTO students (name, marks) VALUES ('Amit', 85)")
cur.execute("INSERT INTO students (name, marks) VALUES ('Priya', 92)")
conn.commit()

cur.execute("SELECT * FROM students")
print("Database Records:")
for row in cur.fetchall():
    print(row)

conn.close()""",
"output": """Database Records:
(1, 'Amit', 85)
(2, 'Priya', 92)""",
"result": "Database integration with a Python backend was successfully implemented using SQLite, with table creation, insertion, and retrieval of records verified.",
},
{
"num": 15, "title": "CRUD Operations in a Full-Stack Web Application",
"aim": "To study and implement Create, Read, Update, and Delete (CRUD) operations on a database using Python and SQLite.",
"theory": "CRUD refers to the four fundamental operations performed on persistent data in virtually any data-driven application: Create (insert new records into the database, corresponding to the SQL INSERT statement and typically an HTTP POST request in a web API), Read (retrieve one or more existing records, corresponding to SQL SELECT and an HTTP GET request), Update (modify the values of existing records, corresponding to SQL UPDATE and an HTTP PUT/PATCH request), and Delete (remove records, corresponding to SQL DELETE and an HTTP DELETE request). These four operations form the backbone of full-stack web applications: the frontend (UI) captures user actions, the backend (server-side application logic, e.g., Flask or Express) receives requests and translates them into the appropriate database operations, and the database (here, SQLite) persists the data using SQL statements executed through a database cursor, with every write operation (Create/Update/Delete) requiring an explicit commit() to make the change durable. Understanding CRUD is essential because nearly every real-world feature - user registration, a shopping cart, a content management system - is ultimately composed of these same four operations applied to different types of data.",
"algorithm": [
"Connect to an SQLite in-memory database and create a products table.",
"Create: insert new product records and commit the transaction, then display all records.",
"Read: query and display a specific product record by its id.",
"Update: modify the price of a specific product and commit, then display all records.",
"Delete: remove a specific product record and commit, then display the remaining records.",
"Close the database connection.",
],
"code": """import sqlite3

conn = sqlite3.connect(":memory:")
cur = conn.cursor()
cur.execute("CREATE TABLE products (id INTEGER PRIMARY KEY, name TEXT, price INTEGER)")

# Create
cur.execute("INSERT INTO products (name, price) VALUES ('Pen', 10)")
cur.execute("INSERT INTO products (name, price) VALUES ('Book', 50)")
conn.commit()
print("After CREATE:")
for row in cur.execute("SELECT * FROM products"):
    print(row)

# Read
print("\\nREAD product id=1:")
print(cur.execute("SELECT * FROM products WHERE id=1").fetchone())

# Update
cur.execute("UPDATE products SET price=15 WHERE name='Pen'")
conn.commit()
print("\\nAfter UPDATE:")
for row in cur.execute("SELECT * FROM products"):
    print(row)

# Delete
cur.execute("DELETE FROM products WHERE name='Book'")
conn.commit()
print("\\nAfter DELETE:")
for row in cur.execute("SELECT * FROM products"):
    print(row)

conn.close()""",
"output": """After CREATE:
(1, 'Pen', 10)
(2, 'Book', 50)

READ product id=1:
(1, 'Pen', 10)

After UPDATE:
(1, 'Pen', 15)
(2, 'Book', 50)

After DELETE:
(1, 'Pen', 15)""",
"result": "CRUD operations (Create, Read, Update, Delete) were successfully implemented and verified on an SQLite database, forming the core of a full-stack web application's backend logic.",
},
]


def add_heading(doc, text):
    h = doc.add_heading(text, level=2)
    return h


def add_code_block(doc, code):
    p = doc.add_paragraph()
    run = p.add_run(code)
    run.font.name = "Consolas"
    run.font.size = Pt(10)
    p.paragraph_format.left_indent = Pt(12)


def build_docx(practical, out_path):
    doc = Document()
    title = doc.add_heading(f"Practical {practical['num']}: {practical['title']}", level=1)

    add_heading(doc, "Aim")
    doc.add_paragraph(practical["aim"])

    add_heading(doc, "Theory")
    doc.add_paragraph(practical["theory"])

    if practical.get("algorithm"):
        add_heading(doc, "Algorithm")
        for step in practical["algorithm"]:
            doc.add_paragraph(step, style="List Number")

    add_heading(doc, "Program")
    add_code_block(doc, practical["code"])

    add_heading(doc, "Output")
    add_code_block(doc, practical["output"])

    add_heading(doc, "Result / Conclusion")
    doc.add_paragraph(practical["result"])

    doc.save(out_path)


def build_combined_pdf(practicals, out_path, heading_text):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.lib.enums import TA_LEFT
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, PageBreak, Preformatted,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "PracticalTitle", parent=styles["Heading1"], fontSize=15, spaceAfter=10,
    )
    section_style = ParagraphStyle(
        "SectionHeading", parent=styles["Heading2"], fontSize=12,
        spaceBefore=10, spaceAfter=4, textColor="#1a1a1a",
    )
    body_style = ParagraphStyle(
        "Body", parent=styles["Normal"], fontSize=10, leading=14, alignment=TA_LEFT,
    )
    code_style = ParagraphStyle(
        "Code", parent=styles["Code"], fontSize=8.5, leading=11,
        backColor="#f4f4f4", borderPadding=6,
    )

    doc = SimpleDocTemplate(
        out_path, pagesize=A4,
        leftMargin=2 * cm, rightMargin=2 * cm, topMargin=2 * cm, bottomMargin=2 * cm,
    )

    story = [Paragraph(heading_text, title_style), Spacer(1, 8)]

    for i, p in enumerate(practicals):
        story.append(Paragraph(f"Practical {p['num']}: {p['title']}", title_style))

        story.append(Paragraph("Aim", section_style))
        story.append(Paragraph(p["aim"], body_style))

        story.append(Paragraph("Theory", section_style))
        story.append(Paragraph(p["theory"], body_style))

        if p.get("algorithm"):
            story.append(Paragraph("Algorithm", section_style))
            for j, step in enumerate(p["algorithm"], 1):
                story.append(Paragraph(f"{j}. {step}", body_style))

        story.append(Paragraph("Program", section_style))
        story.append(Preformatted(p["code"], code_style))

        story.append(Paragraph("Output", section_style))
        story.append(Preformatted(p["output"], code_style))

        story.append(Paragraph("Result / Conclusion", section_style))
        story.append(Paragraph(p["result"], body_style))

        if i != len(practicals) - 1:
            story.append(PageBreak())

    doc.build(story)


def main():
    for p in PRACTICALS:
        folder_name = f"Practical_{p['num']:02d}"
        folder_path = os.path.join(ROOT, folder_name)
        os.makedirs(folder_path, exist_ok=True)
        out_path = os.path.join(folder_path, f"{folder_name}.docx")
        build_docx(p, out_path)
        print(f"Created {out_path}")

    top5 = [p for p in PRACTICALS if p["num"] <= 5]
    combined_out = os.path.join(ROOT, "Practicals_PDF", "Practicals_01-05_Combined.pdf")
    build_combined_pdf(top5, combined_out, "Practicals 1-5: Combined Report")
    print(f"Created {combined_out}")


if __name__ == "__main__":
    main()
