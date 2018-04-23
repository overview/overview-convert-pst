This is an [Overview](https://github.com/overview/overview-server) converter.
It extracts the contents of `.pst` email archives to `.eml` files.

What it does
============

Currently, we support Microsoft Outlook `.pst` files. The converter:

1. Writes stdin to `input.pst` in the CWD
2. Extracts it, streaming `.eml` files in `multipart/form-data` on stdout

To be extra-clear, the output is a MIME multipart message composed of MIME
multipart messages:

* Each output `.eml` file is a `message/rfc822`. It is usually a MIME multipart
  message: a sequence of "parts" separated by `--RANDOM-BOUNDARY` boundaries.
* This program's output is a _sequence_ of `.eml` files, and that _sequence_
  is encoded as `multipart/form-data`, which is a MIME multipart message
  separated by `--DIFFERENT-RANDOM-BOUNDARY` boundaries.

The output may look something like:

```
HTTP message
╟─001.eml
║ ╟─body, encoded as text/plain
║ ╟─body, encoded as text/html
║ ╟─body, encoded as application/rtf
║ ╟─attachment1.doc
║ ╙─attachment2.pdf
╟─002.eml
║ ╟─body, encoded as text/plain
║ ╟─body, encoded as text/html
║ ╟─body, encoded as application/rtf
║ ╙─attachment3.pdf
╟─003.eml
║ ╙─body, encoded as text/plain
...
```

PST files store messages in a proprietary format. This converter encodes each as
a `message/rfc822` message.

Usage
=====

In an Overview cluster, you'll want to use the Docker container:

`docker run -e POLL_URL=http://worker-url:9032/Pst overview/overview-convert-pst:0.0.1`

Developing
==========

`./dev` will connect to the `overviewserver_default` network and run with
`POLL_URL=http://overview-worker:9032/Pst`.

`docker build .` will run tests.

Design decisions
----------------

`src/extract-pst.c` is in C. That's because we rely on
[libpst](http://www.five-ten-sg.com/libpst/) to do the extraction: its Python
wrapper is long obsolete and its `readpst` executable doesn't do streaming or
progress reports.

This extractor output is simple: `.eml` files. That's because:

* Overview's `.eml` converter can be reused to convert `eml` and `mbox` files.
* Overview can let users download the individual `.eml` files to open them in
  virtually any email-related program.
* Overview can convert `.eml` files in parallel.

License
-------

This software is Copyright 2011-2018 Jonathan Stray, and distributed under the
terms of the GNU Affero General Public License. See the LICENSE file for details.
