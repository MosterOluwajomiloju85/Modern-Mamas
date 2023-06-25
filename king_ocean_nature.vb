Option Explicit

Dim strBlogTitle As String

Dim strLatestPost As String

Dim strAuthor As String

Dim strLatestComment As String

Dim strCategories As String

Dim strTags As String

Dim strArchive As String

Dim strRecentPost As String

Sub Main()

'Set the blog title

strBlogTitle = "A Parenting and Lifestyle Blog"

'Set the latest post

strLatestPost = "Happy First Birthday to Little Jack!"

'Set the author of the blog

strAuthor = "Elizabeth Johnson"

'Set the latest comment on the blog

strLatestComment = "This post has really been helpful! Thanks for sharing!"

'Set the categories

strCategories = "Parenting Tips, Recipes, Health & Wellness, Lifestyle"

'Set the tags

strTags = "Mommy-Daughter Bonding, Newborn Essentials, Low-Fat Recipes, Home Decor, Self-Care"

'Set the archive

strArchive = "May-June 2018, July-August 2018, September-October 2018, November-December 2018"

'Set the recent post

strRecentPost = "What I Learned From Being a First-Time Mom"

'Display the blog

Call DisplayBlog

End Sub

Sub DisplayBlog()

'Display the blog title

Print strBlogTitle

'Display the latest post

Print strLatestPost

'Display the author

Print "by " & strAuthor

'Display the latest comment

Print strLatestComment

'Display the categories

Print "Categories: " & strCategories

'Display the tags

Print "Tags: " & strTags

'Display the archive

Print "Archive: " & strArchive

'Display the recent post

Print "Recent Post: " & strRecentPost

End Sub