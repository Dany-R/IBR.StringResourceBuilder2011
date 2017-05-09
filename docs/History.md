# History

## The Beginning

In 2006 I created the first version of this extension with VS2005 'afoot' (no support tool like VSPackage Builder).  It had almost the same functionality as the current project.

I have been in dire need of such a tool because at work I needed to translate a growing C# application and I have been tired of creating the string resources manually (it is astounding how many texts an application 'grows').  So I decided (for coding is also a hobby of mine) to make an extension for VS to make my life more comfortable.

I used this extension for quite some time until we upgraded to VS2010 at work.  The extension still worked then.  But then around 2010 I found the VSPackage Builder when I thought about using it.  In lack of a better idea I just re-created the String Resource Builder with said extension.

I wanted to contribute the project to the community for a while now but didn't consider it 'ripe', so I shelved this problem.  But I made no changes anymore so eventually I decided to publish in the VisualStudio Gallery and here at CodePlex.  Maybe one will find it useful.

Between 2012-06-09 and 2012-07-01 I had 116 downloads (some of course where my own) but no review, rating or Q&A yet.  I found some bugs in that time and fixed them, because I'm still using it productive at work of course.

In November 2013 I reached the 1000th download and four reviews, though not all of them seeming fair :P.

In February 2015 I reached the 2000th download but only one additional review.

In January 2017 I reached the 3000th download and four additional reviews, now only supporting VS2015 and VS2017.

## The Present

It's pretty quiet regarding my private projects at the moment. :)

I just moved from CodePlex to GitHub because it will be closed down for active contributions in 2017.

## Obstacles

There were some problems concerning the work with VSPackage Builder and the project name I have chosen.  It is per se not working with names containing a dot (I know that there are at least two opinions regarding this, which I don't want to discuss here, please).  I have patched some of the text templating (.tt) files.

There also is the lack of support of PNG images for buttons in VSPackage Builder.  At the moment only BMP images are allowed.  I patched this as well but this is a bit awkward because for the designer I have to provide the images as BMP as well (that's why they are also in version control).

I'm truly sad to say that there seems to be an end to the VSPackage Builder development, since the last update has been 2010-09-04 and there are about no answers to questions anymore.  I really like that extension though.

As of V1.5 R2 (20) I removed VSPackage Builder code generation as it is no longer maintained.
