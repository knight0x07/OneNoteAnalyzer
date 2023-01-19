# OneNoteAnalyzer
A C# based tool for analyzing malicious OneNote documents

## Description

Recently we came across few malicious OneNote Documents been distributed in-the-wild by various threat actors. This gave us an idea to develop "OneNoteAnalyzer" which would help in analysing such malicious OneNote documents without executing them. Now lets take a look at the features that the tool offers.

## Features

After providing the file path of the Malicious OneNote document. The OneNoteAnalyzer extracts:
- Attachments from OneNote Document along with the Actual Attachment Path, Filename and size
- Page MetaData from OneNote Document - Title, Author, CreationTime, LastModifiedTime
- Images from OneNote Document along with the HyperLink URLs if any
- Pagewise Text from OneNote Document
- HyperLinks from OneNote Document along with the overlay text
- and Converts OneNote Document to Image

## Usage

![usage](https://user-images.githubusercontent.com/60843949/212981619-8fe1c6c0-0ffb-4e37-869c-0febbd1484ea.PNG)

## Demonstration

In order to execute OneNoteAnalyzer against malicious OneNote Documents we provide the path of the OneNote Document as shown below.

<img width="1000" alt="1" src="https://user-images.githubusercontent.com/60843949/212985874-f306bad8-d348-413c-8422-e191caced924.PNG">

Upon execution OneNoteAnalyzer extracts the Attachments from the OneNoteDocument in the "OneNoteAttachments" folder. Here the Actual Attachment path i.e the path from where the attachment was been uploaded can be seen in the console along with the extracted filename and size of the attachment.

![2](https://user-images.githubusercontent.com/60843949/212986509-fc825b7f-312b-48b3-b867-ddb5d3b2d58a.PNG)
![2](https://user-images.githubusercontent.com/60843949/212986244-efd3b12b-5864-4971-93ae-8cecbd3a74ce.PNG)

OneNote Attachments extracted in the OneNoteAttachments Folder:

![12](https://user-images.githubusercontent.com/60843949/212985277-d4199cda-272e-4917-974a-89908c57c4f4.PNG)
![10](https://user-images.githubusercontent.com/60843949/212986790-6dcb1b7f-be2e-4a85-9c6f-26a1281b37e5.PNG)

Next it extracts the Pagewise Metadata from the OneNote Document as shown below.

<img width="700" alt="3" src="https://user-images.githubusercontent.com/60843949/212987251-1b787542-ff47-46c5-8f0f-5bba024318f2.PNG">

Then it also extracts all the images in the OneNote Document as shown below:

<img width="700" alt="4" src="https://user-images.githubusercontent.com/60843949/212987927-e45c33bd-7470-4319-9f0e-f3e992651595.PNG">
<img width="700" alt="4" src="https://user-images.githubusercontent.com/60843949/212988364-c704a84e-33e4-4a89-8e78-73c0bfb50cf1.PNG">

The extracted images are been saved in the OneNoteImages folder as shown below.

<img width="700" alt="11" src="https://user-images.githubusercontent.com/60843949/212988108-ea337afc-da10-478f-945c-70c9b22cf56e.PNG">
<img width="500" alt="9" src="https://user-images.githubusercontent.com/60843949/212988414-c7af58ce-03f8-43ae-8704-3614cd545196.PNG">

Further the tool extracts Pagewise Text from the OneNote Document 

<img width="2000" alt="5" src="https://user-images.githubusercontent.com/60843949/212988882-b4f9d1d6-30f4-4142-98ff-16d46c2e3347.PNG">

and saves it in the OneNoteText Folder as shown in the screenshot below

![10](https://user-images.githubusercontent.com/60843949/212989015-7ad44df8-cc7a-4550-a60c-7e969ce39904.PNG)

Addtionally it extracts HyperLinks from OneNote Document along with the overlay text as shown in the screenshot below.

![1](https://user-images.githubusercontent.com/60843949/212989534-fe9a6ebe-540a-4e91-b58f-f2a5089078a3.PNG)

The extracted Hyperlinks are stored in the OneNoteHyperLinks Folder - onenote_hyperlinks.txt

![hyperlinks](https://user-images.githubusercontent.com/60843949/213367331-7289c92f-4198-436d-8054-a3799c0f13b5.PNG)

Finally the tool converts the OneNoteDocument into an Image and saves it shown in the following manner.

<img width="800" alt="7" src="https://user-images.githubusercontent.com/60843949/212990203-9c836f27-088a-4038-ae1e-e7be62b2e6bd.PNG">

Saved Image-1:

<img width="577" alt="8" src="https://user-images.githubusercontent.com/60843949/212990317-965fb1fb-d2f2-4c56-996e-1b20bf879b11.PNG">

Saved Image-2:

<img width="483" alt="8" src="https://user-images.githubusercontent.com/60843949/212991656-58da71e6-cfe4-4441-aaa9-6d55a97cb8fc.PNG">

Once the execution is completed the extracted data is been stored in an Export Directory "OneNoteFilename_content" in the current working directory as seen in the screenshot below

![export_1](https://user-images.githubusercontent.com/60843949/213369569-57913f05-5809-440c-bd3f-9bd9ede8b9b6.PNG)

## Setup Information

- Copy "Program.cs" in Visual Studio
- Install "Aspose.Note 18.1.0" from Nuget Packages
- Build the project!

## Updates
- Added Export Directory where all the extracted data from the OneNote Document is been dumped (Compiled binary can be downloaded from Releases)

## References

https://docs.aspose.com/note/net

Thankyou! =)







































