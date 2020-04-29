# DayBook-XL

Previously Expense Sheet Generator.

A simple CLI app to generate expense sheet for a friend.

## History

One of friend self employed and has to report year end income and expense details. Idea he has is create an excel sheet. The sheet is grouped by months with total of items (income/expense).
Items are horizontally aligned in columns with respective row as date for each month.

## Problem

Initially, he had income/expense small list. To solve it created a simple excel template which he can replicate each year end. However the list started to be bigger and bigger so the calculations started to get messed up. Each year either date messed up. Sometime, total for each month get messed. Sometime, day total at the end of the row and even total at the bottom.

## Rejected Thought

Thought of creating a web app or an electron app which takes each entry and generate excel sheet ready to be send to an account. However, that would require resources and overkill so got rejected.

## Solution

Just automate my task. It would have solved the issue if just automate my task with any list of items and generate excel work sheet with exact same features. To simplify and used just by me every year, I created this cli app that simply generates expected excel file. I will amend item list and run script to generate an expected worksheet :)
