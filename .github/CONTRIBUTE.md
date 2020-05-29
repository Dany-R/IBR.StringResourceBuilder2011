[Home][home]

# Contributing

When contributing to this repository, please first discuss the change you wish to make via issue,
email, or any other method with the owners of this repository before making a change.

In the Web there are a lot of useful blogs which describe how contributions could work, like this
one (partially used in this document): 
[https://opensource.guide/how-to-contribute/#how-to-submit-a-contribution][how-to-submit-a-contribution]
(licensed under [CC BY 4.0][CC]).

## Code of conduct

Please note we have a [code of conduct](CODE_OF_CONDUCT.md), please follow it in all your
interactions with the project.

## Using the GitHub issue tracker

The GitHub issue tracker is the preferred channel for:

* [bug reports and features requests (issues)](#issues)
* [pull requests](#pull-requests)

But please respect the following restrictions:

* Please **do not** use the issue tracker for personal support requests.  Stack Overflow is a better
  place to get help.
* Please **do not** derail or troll issues. Keep the discussion on topic and respect the opinions of
  others (see also [code of conduct](CODE_OF_CONDUCT.md)).

**Give context.** Help others get quickly up to speed. If you’re running into an error, explain what
you’re trying to do and how to reproduce it. If you’re suggesting a new idea, explain why you think
it’d be useful to the project (not just to you!).

### <a name="issues"></a> Bug reports and feature requests (issues)

You should usually open an issue in the following situations:

* Report an error you can’t solve yourself
* Discuss a high-level topic or idea (ex. community, vision, policies)
* Propose a new feature or other project idea

Tips for communicating on issues:

* **If you see an open issue that you want to tackle**, comment on the issue to let people know
  you’re on it. That way, people are less likely to duplicate your work.
* **If an issue was opened a while ago**, it’s possible that it’s being addressed somewhere else,
  or has already been resolved, so comment to ask for confirmation before starting work.
* **If you opened an issue**, but figured out the answer later on your own, comment on the issue to
  let people know, then close the issue. Even documenting that outcome is a contribution to the
  project.

### <a name="pull-requests"></a> Pull requests (PR)

You should usually open a pull request in the following situations:

* Submit trivial fixes (ex. a typo, broken link, or obvious error)
* Start work on a contribution that was already asked for, or that you’ve already discussed, in an
  issue

A pull request doesn’t have to represent finished work. It’s usually better to open a pull request
early on, so others can watch or give feedback on your progress. Just mark it as a “WIP” (Work in
Progress) in the subject line. You can always add more commits later.

For those who are new to GitHub, here’s how to submit a pull request:

* [**Fork the repository**][fork] and clone it locally. Connect your local to the original “upstream”
  repository by adding it as a remote. Pull in changes from “upstream” often so that you stay up to
  date so that when you submit your pull request, merge conflicts will be less likely. (See more
  detailed instructions [here][forksync].)
* [**Create a branch**][branch] for your edits.
* **Reference any relevant issues** or supporting documentation in your PR (ex. “Closes #37.”)
* **Include screenshots of the before and after** if your changes include differences in the GUI.
  Drag and drop the images into the body of your pull request.
* **Test** your changes! Run your changes against any existing tests if they exist and create new
  ones when needed. Whether tests exist or not, make sure your changes don’t break the existing
  project. _(no automated tests yet so please test it manually as bes as you could)_
* **Contribute** in the style of the project to the best of your abilities (see
  [Code guidelines](#code-guidelines)). This may mean using indents, semi-colons or comments
  differently than you would in your own repository, but makes it easier for the maintainer to merge,
  others to understand and maintain in the future.

If this is your first pull request, check out [Make a Pull Request][makepull], which **@kentcdodds**
created as a free walkthrough resource.

## <a name="code-guidelines"></a> Code guidelines

Starting with VS2017, the most basic code guidelines will be enforced (or suggested) by an
`.editorconfig` file which is integrated with the solution.


[home]: https://github.com/Dany-R/IBR.StringResourceBuilder2011
[how-to-submit-a-contribution]: https://opensource.guide/how-to-contribute/#how-to-submit-a-contribution
[CC]: https://creativecommons.org/licenses/by/4.0/
[fork]: https://guides.github.com/activities/forking/
[forksync]: https://help.github.com/articles/syncing-a-fork/
[branch]: https://guides.github.com/introduction/flow/
[makepull]: http://makeapullrequest.com/