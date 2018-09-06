# Automated Marketing Dashboard in InDesign with Python
Like all good marketers, we had a simple challenge: prove that our marketing efforts were bearing fruit. The challenge was that the dashboard we had created in InDesign took hours to calculate and create.

Luckily, someone mentioned that Python should be able to do something like that automatically.

The .py file included is my solution. Using pandas and matplotlib, it takes our monthly marketing data, runs calculations, and creates all of the charts we need as EPS files. (Shout out to matplotlib's Wedges and Circles!)

It then creates an InDesign document, sets some brand guideline-compliant styles, then lays out all of the charts and data in an easy-to-read way. It does not currently save the file, but that's to encourage the team member overseeing the creation of this dashboard to make sure that everything is in the correct place and looks good before proceeding.

This was created as I was learning Python, so please be nice and/or let me know how I can improve my code.
