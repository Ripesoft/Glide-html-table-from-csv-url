# Glide Yes-Code Template

This is a sample and template for implementing your own Yes-Code column.  It's a fully functional implementation of a simple substring column.

## Getting started

The first thing to do is to fork [this repl](https://replit.com/@MarkProbst/glide-yes-code-substring#README.md) so you have your own copy of it that you can modify.  You can then try the URL of your copy (you should see it in the mini-browser on your right) in Glide.  Just make sure to delete any columns you made with it, and reload Glide before you move on.

To implement your own column you only have to change two files:

* `manifest.json` contains metadata about your function, such as its name, a description, what parameters it takes, and what kind of result it produces.
* `function.js` is the actual JavaScript code of your column.

The manifest should be mostly self-explanatory, except for what `type`s are allowed, which at the moment is only `string`, `number`, and `boolean`.  Glide will refuse to load ill-formed manifests.  If you make a change to your manifest you'll have to reload Glide - it won't reload it on its own.  We'll fix that.

The implementation file has comments explaining it.
