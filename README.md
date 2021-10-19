# Robot. Homework 3
***
The link to loom video *https://www.loom.com/share/ba656ddf840f406da60d4619fed6f069*
***
The main goal of the robot is to get the information about the articles on the given theme from the  *https://www.semanticscholar.org/*
***
### The robot functions
* #### Search articles by specific topic on N pages
* #### Get title, author, source, description, number of citations, article file (if available).

### The */conf.py* is used to configure robot
* #### query - The theme of the articles to search
* #### num_page - The number of pages
* #### receiver - The email of the receiver
* #### sender_email - The email of the sender ( Is empty in repository for security reasons )
* #### sender_password - The password of the sender ( Is empty in repository for security reasons )

### The */main.py* is the main file of robot, which is used to run it.