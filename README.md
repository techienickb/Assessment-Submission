## Assessment Submissions

![Screenshot](/Assessment-Submission/Annotation2020-05-28152213.png "Screenshot")

This webpart is available in the Cardiff University SharePoint App Store, see Using the Assessment Submissions Form section for instructions.

There is also a webpart for monitoring folders for submissions.

The sharepoint site you are adding this app to should not be accessible to students, students will get access to just their folder via a sharing link

When you complete the form it will:
- Creates a folder called Submissions in the root of your sharepoint site
- Creates a folder named as your submission
- Student folders are then created based on your criteria for folder naming
- A sharing link is then sent to that students so they can access the folder


To remove permissions click, select the Submission Folder and click remove, it will:
- Go through each folder and delete any sharing link permissions on that folder

Note again: Students should not have access to this sharepoint site.  You can use a private channel inside teams if you want to keep submissions inside easily accessibly inside a class team.


### Using the Assessment Submissions Form

1. In Microsoft Teams click on the Files tab and select Open in SharePoint

![Screenshot of Teams](/Assessment-Submission/Annotation2020-05-28163932.png "Screenshot of Teams")

2. Click on Site contents on the left

![Screenshot of Site contents link](/Assessment-Submission/Annotation2020-05-28164021.png "Screenshot of Site contents link")

3. Click on New and select App

![Screenshot of New > Select App](/Assessment-Submission/Annotation2020-05-28164127.png "Screenshot of New > Select App")

4. Select From your Organization from the left
5. Click on Assessment Submission

![Screenshot of Assessment Submission App](/Assessment-Submission/Annotation2020-05-28164327.png "Screenshot of Assessment Submission App")

6. Click on New and select Page

![Screenshot of New > Page](/Assessment-Submission/Annotation2020-05-28164441.png "Screenshot of New > Page")

7. Enter a page name, this can be anything you want

![Screenshot of naming the page](/Assessment-Submission/Annotation2020-05-28165907.png "Screenshot of naming the page")

8. Click on the + in the middle of the page

![Screenshot of adding a web part 1](/Assessment-Submission/Annotation2020-05-28164539.png "Screenshot of adding a web part 1")

8. Search for Assessment - add the Form.  You can also add the Monitor if you wish.

![Screenshot of adding a web part 2](/Assessment-Submission/Annotation2020-05-28170002.png "Screenshot of adding a web part 2")

9. Publish the Page







### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Build options

gulp clean - Clears the build folders
gulp test - Tests the build
gulp serve - Starts a basic web server to serve the webpart(s) during debug testing in the work bench
gulp bundle - Compiles a the typescript into a javascript bundle
gulp package-solution - Packages the bundle into a spfx package