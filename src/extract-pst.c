/***
 * readpst.c
 * Part of the LibPST project
 * Written by David Smith
 *            dave.s@earthcorp.com
 */

#include <ctype.h>
#include <errno.h>
#include <math.h>
#include <stdio.h>
#include <string.h>
#include <inttypes.h>

// libpst includes:
#include <libpst.h>
#include <libstrfunc.h>
#include <lzfu.h>
#include <timeconv.h>

// max size of the c_time char*. It will store the date of the email
#define C_TIME_SIZE 500

#define DEBUG_ENT(x)
#define DEBUG_INFO(x)
#define DEBUG_WARN(x)
#define DEBUG_RET(x)

typedef struct {
    size_t n_processed;
    size_t n_total;
} Progress;

size_t    process(pst_file *pstfile, pst_item *outeritem, pst_desc_tree *d_ptr, size_t starting_index, const char* outer_name, Progress* progress);
void      removeCR(char *c);
void      usage();
char*     my_stristr(char *haystack, char *needle);
void      write_embedded_message(pst_item_attach* attach, char *boundary, pst_file* pstfile, char** extra_mime_headers);
void      write_inline_attachment(pst_item_attach* attach, char *boundary, pst_file* pst);
int       valid_headers(char *header);
void      header_has_field(char *header, char *field, int *flag);
void      header_get_subfield(char *field, const char *subfield, char *body_subfield, size_t size_subfield);
char*     header_get_field(char *header, char *field);
char*     header_end_field(char *field);
void      header_strip_field(char *header, char *field);
int       test_base64(char *body, size_t len);
void      find_rfc822_headers(char** extra_mime_headers);
void      write_body_part(pst_string *body, char *mime, char *charset, char *boundary, pst_file* pst);
void      write_schedule_part_data(pst_item* item, const char* sender, const char* method);
void      write_schedule_part(pst_item* item, const char* sender, const char* boundary);
void      write_normal_email(pst_item* item, pst_file* pst, int embedding, char** extra_mime_headers);
void      write_vcard(pst_item *item, pst_item_contact* contact, char comment[]);
int       write_extra_categories(pst_item* item);
void      write_journal(pst_item* item);
void      write_appointment(pst_item *item);
char*     quote_string(char *inp);

const char*  prog_name;
static size_t MIN_N_DIGITS = 4;
#define MIN_N_DIGITS_S "4"

// default mime-type for attachments that have a null mime-type
#define MIME_TYPE_DEFAULT "application/octet-stream"
#define RFC822            "message/rfc822"

// Output type mode flags
#define OTMODE_EMAIL        1
#define OTMODE_APPOINTMENT  2
#define OTMODE_JOURNAL      4
#define OTMODE_CONTACT      8

// output settings for RTF bodies
// filename for the attachment
#define RTF_ATTACH_NAME "rtf-body.rtf"
// mime type for the attachment
#define RTF_ATTACH_TYPE "application/rtf"

const char* mime_boundary;
const char* json_template;

void
die(const char* message)
{
	printf(
		"\r\n--%s\r\nContent-Disposition: form-data; name=error\r\n\r\n%s\r\n--%s--",
		mime_boundary,
		message,
		mime_boundary
	);
	exit(0);
}

void
output_part(const char* name, const char* body)
{
	printf(
		"\r\n--%s\r\nContent-Disposition: form-data; name=%s\r\n\r\n%s",
		mime_boundary,
		name,
		body
	);
}

void
output_indexed_part(int index, const char* ext, const char* body)
{
	printf(
		"\r\n--%s\r\nContent-Disposition: form-data; name=%d%s\r\n\r\n%s",
		mime_boundary,
		index,
		ext,
		body
	);
}

void
output_json(int index, const char* filename)
{
	const char* filename_pos = strstr(json_template, "FILENAME\"");
	if (!filename_pos) {
		die("Expected placeholder 'FILENAME' to exist in JSON template");
	}

	output_indexed_part(index, ".json", "");
	fwrite(json_template, sizeof(char), filename_pos - json_template, stdout);
	fputs(filename, stdout);
	fputs(filename_pos + strlen("FILENAME"), stdout);
}

void
increment_and_output_progress(Progress* progress)
{
    progress->n_processed += 1;
	printf(
		"\r\n--%s\r\nContent-Disposition: form-data; name=progress\r\n\r\n{\"children\":{\"nProcessed\":%d,\"nTotal\":%d}}",
		mime_boundary,
		progress->n_processed,
		progress->n_total
	);
}

void
output_json_and_blob_and_progress(int index, Progress* progress, const char* filename, const char* body)
{
	output_json(index, filename);
	output_indexed_part(index, ".blob", body);
	increment_and_output_progress(progress);
}

void
output_appointment(int index, Progress* progress, const char* filename, pst_item* item)
{
    output_json(index, filename);
    output_indexed_part(index, ".blob", "");
    write_appointment(item);
    increment_and_output_progress(progress);
}

void
output_journal(int index, Progress* progress, const char* filename, pst_item* item)
{
    output_json(index, filename);
    output_indexed_part(index, ".blob", "");
    write_journal(item);
    increment_and_output_progress(progress);
}

void
output_vcard(int index, Progress* progress, const char* filename, pst_item* item)
{
    output_json(index, filename);
    output_indexed_part(index, ".blob", "");
    write_vcard(item, item->contact, item->comment.str);
    increment_and_output_progress(progress);
}

void
output_email(int index, Progress* progress, const char* filename, pst_item* item, pst_file* pstfile)
{
    output_json(index, filename);
    output_indexed_part(index, ".blob", "");
    char *extra_mime_headers = NULL;
    write_normal_email(item, pstfile, 0, &extra_mime_headers);
    increment_and_output_progress(progress);
}

void*
malloc_or_die(size_t size)
{
	void* ret = malloc(size);
	if (ret == NULL) {
		die("out of memory because a message was too large");
	}
	return ret;
}

char*
strdup_or_die(const char* s)
{
	char* ret = strdup(s);
	if (ret == NULL) {
		die("out of memory because a string was too large");
	}
	return ret;
}

char*
strdup_parent_sep_child_or_die(const char* parent, const char* sep, const char* child)
{
    size_t len = strlen(parent) + strlen(child) + strlen(sep) + 1;
    char* ret = malloc_or_die(len);
    snprintf(ret, len, "%s%s%s", parent, sep, child);
    return ret;
}

char*
strdup_parent_slash_num_dot_or_die(const char* parent, size_t n, const char* ext)
{
    size_t n_digits = (int) ceil(log10(n));
    if (n_digits < MIN_N_DIGITS) {
        n_digits = MIN_N_DIGITS;
    }
    size_t len = strlen(parent) + strlen("/") + n_digits + strlen(ext) + 1;
    char* ret = malloc_or_die(len);
    snprintf(ret, len, "%s/%0" MIN_N_DIGITS_S "d%s", parent, n, ext);
    ret[len] = '\0';
    return ret;
}

/**
 * Outputs parts for the given item and returns starting_index + (nPartsOutput)
 */
size_t
process(pst_file *pstfile, pst_item *outeritem, pst_desc_tree *d_ptr, size_t starting_index, const char* outer_name, Progress* progress)
{
    pst_item *item = NULL;
    size_t item_number = 1;
    size_t index = starting_index;

    DEBUG_ENT("process");

    for (; d_ptr; d_ptr = d_ptr->next) {
        DEBUG_INFO(("New item record\n"));
        if (!d_ptr->desc) {
            DEBUG_WARN(("ERROR item's desc record is NULL\n"));
            continue;
        }
        DEBUG_INFO(("Desc Email ID %#"PRIx64" [d_ptr->d_id = %#"PRIx64"]\n", d_ptr->desc->i_id, d_ptr->d_id));

        item = pst_parse_item(pstfile, d_ptr, NULL);
        DEBUG_INFO(("About to process item\n"));
        pst_convert_utf8(item, &item->file_as);

        if (!item) {
            DEBUG_INFO(("A NULL item was seen\n"));
            continue;
        }

        if (item->subject.str) {
            DEBUG_INFO(("item->subject = %s\n", item->subject.str));
        }

        if (item->folder && item->file_as.str) {
            DEBUG_INFO(("Processing Folder \"%s\"\n", item->file_as.str));
            // If this is a non-empty folder and we have not yet seen a
            // folder, its item_count is the total number of items in the
            // PST.
            if (!progress->n_total && item->folder->item_count) {
                progress->n_total = item->folder->item_count;
            }

            if (d_ptr->child) {
                //if this is a non-empty folder other than deleted items, we want to recurse into it
                char* inner_name = strdup_parent_sep_child_or_die(outer_name, "/", item->file_as.str);
                index = process(pstfile, item, d_ptr->child, index, inner_name, progress);
                free(inner_name);
            }
        } else if (item->contact && (item->type == PST_TYPE_CONTACT)) {
            DEBUG_INFO(("Processing Contact\n"));
            char* inner_name = strdup_parent_slash_num_dot_or_die(outer_name, item_number, ".vcard");
            pst_convert_utf8_null(item, &item->comment);
            output_vcard(index, progress, inner_name, item);
            free(inner_name);
            item_number += 1;
            index += 1;
        } else if (item->email && ((item->type == PST_TYPE_NOTE) || (item->type == PST_TYPE_SCHEDULE) || (item->type == PST_TYPE_REPORT))) {
            DEBUG_INFO(("Processing Email\n"));
            char* inner_name = strdup_parent_slash_num_dot_or_die(outer_name, item_number, ".eml");
            output_email(index, progress, inner_name, item, pstfile);
            free(inner_name);
            item_number += 1;
            index += 1;
        } else if (item->journal && (item->type == PST_TYPE_JOURNAL)) {
            DEBUG_INFO(("Processing Journal Entry\n"));
            char* inner_name = strdup_parent_slash_num_dot_or_die(outer_name, item_number, ".ics");
            output_journal(index, progress, inner_name, item);
            free(inner_name);
            item_number += 1;
            index += 1;
        } else if (item->appointment && (item->type == PST_TYPE_APPOINTMENT)) {
            DEBUG_INFO(("Processing Appointment Entry\n"));
            char* inner_name = strdup_parent_slash_num_dot_or_die(outer_name, item_number, ".ics");
            output_appointment(index, progress, inner_name, item);
            free(inner_name);
            item_number += 1;
            index += 1;
        } else if (item->message_store) {
            // there should only be one message_store, and we have already done it
            DEBUG_WARN(("item with message store content, type %i %s, skipping it\n", item->type, item->ascii_type));
            progress->n_processed += 1;
        } else {
            DEBUG_WARN(("Unknown item type %i (%s) name (%s)\n",
                        item->type, item->ascii_type, item->file_as.str));
            progress->n_processed += 1;
        }
        pst_freeItem(item);
    }
    DEBUG_RET();

    return index;
}


void removeCR (char *c) {
    // converts \r\n to \n
    char *a, *b;
    DEBUG_ENT("removeCR");
    a = b = c;
    while (*a != '\0') {
        *b = *a;
        if (*a != '\r') b++;
        a++;
    }
    *b = '\0';
    DEBUG_RET();
}


char *my_stristr(char *haystack, char *needle) {
    // my_stristr varies from strstr in that its searches are case-insensitive
    char *x=haystack, *y=needle, *z = NULL;
    if (!haystack || !needle) {
        return NULL;
    }
    while (*y != '\0' && *x != '\0') {
        if (tolower(*y) == tolower(*x)) {
            // move y on one
            y++;
            if (!z) {
                z = x; // store first position in haystack where a match is made
            }
        } else {
            y = needle; // reset y to the beginning of the needle
            z = NULL; // reset the haystack storage point
        }
        x++; // advance the search in the haystack
    }
    // If the haystack ended before our search finished, it's not a match.
    if (*y != '\0') return NULL;
    return z;
}


void write_embedded_message(pst_item_attach* attach, char *boundary, pst_file* pstfile, char** extra_mime_headers)
{
    pst_index_ll *ptr;
    DEBUG_ENT("write_embedded_message");
    ptr = pst_getID(pstfile, attach->i_id);

    pst_desc_tree d_ptr;
    d_ptr.d_id        = 0;
    d_ptr.parent_d_id = 0;
    d_ptr.assoc_tree  = NULL;
    d_ptr.desc        = ptr;
    d_ptr.no_child    = 0;
    d_ptr.prev        = NULL;
    d_ptr.next        = NULL;
    d_ptr.parent      = NULL;
    d_ptr.child       = NULL;
    d_ptr.child_tail  = NULL;

    pst_item *item = pst_parse_item(pstfile, &d_ptr, attach->id2_head);
    // It appears that if the embedded message contains an appointment/
    // calendar item, pst_parse_item returns NULL due to the presence of
    // an unexpected reference type of 0x1048, which seems to represent
    // an array of GUIDs representing a CLSID. It's likely that this is
    // a reference to an internal Outlook COM class.
    //      Log the skipped item and continue on.
    if (!item) {
        DEBUG_WARN(("write_embedded_message: pst_parse_item was unable to parse the embedded message in attachment ID %llu", attach->i_id));
    } else {
        if (!item->email) {
            DEBUG_WARN(("write_embedded_message: pst_parse_item returned type %d, not an email message", item->type));
        } else {
            printf("\n--%s\n", boundary);
            printf("Content-Type: %s\n\n", attach->mimetype.str);
            write_normal_email(item, pstfile, 1, extra_mime_headers);
        }
        pst_freeItem(item);
    }

    DEBUG_RET();
}

/**
 * Backslash-escape quotes and backslashes in the given string.
 */
char *quote_string(char *inp) {
    int i = 0;
    int count = 0;
    char *curr = inp;
    while (*curr) {
        *curr++;
        if (*curr == '\"' || *curr == '\\') {
            count++;
        }
        i++;
    }
    char *res = malloc(i + count + 1);
    char *curr_in = inp;
    char *curr_out = res;
    while (*curr_in) {
        if (*curr_in == '\"' || *curr_in == '\\') {
            *curr_out++ = '\\';
        }
        *curr_out++ = *curr_in++;
    }
    *curr_out = '\0';
    return res;
}

void write_inline_attachment(pst_item_attach* attach, char *boundary, pst_file* pst)
{
    DEBUG_ENT("write_inline_attachment");
    DEBUG_INFO(("Attachment Size is %#"PRIx64", data = %#"PRIxPTR", id %#"PRIx64"\n", (uint64_t)attach->data.size, attach->data.data, attach->i_id));

    if (!attach->data.data) {
        // make sure we can fetch data from the id
        pst_index_ll *ptr = pst_getID(pst, attach->i_id);
        if (!ptr) {
            DEBUG_WARN(("Couldn't find ID pointer. Cannot save attachment to file\n"));
            DEBUG_RET();
            return;
        }
    }

    printf("\n--%s\n", boundary);
    if (!attach->mimetype.str) {
        printf("Content-Type: %s\n", MIME_TYPE_DEFAULT);
    } else {
        printf("Content-Type: %s\n", attach->mimetype.str);
    }
    printf("Content-Transfer-Encoding: base64\n");

    if (attach->content_id.str) {
        printf("Content-ID: <%s>\n", attach->content_id.str);
    }

    if (attach->filename2.str) {
        // use the long filename, converted to proper encoding if needed.
        // it is already utf8
        char *escaped = quote_string(attach->filename2.str);
        pst_rfc2231(&attach->filename2);
        printf("Content-Disposition: attachment; \n        filename*=%s;\n", attach->filename2.str);
        // Also include the (escaped) utf8 filename in the 'filename' header directly - this is not strictly valid
        // (since this header should be ASCII) but is almost always handled correctly (and in fact this is the only
        // way to get MS Outlook to correctly read a UTF8 filename, AFAICT, which is why we're doing it).
        printf("        filename=\"%s\"\n\n", escaped);
        free(escaped);
    }
    else if (attach->filename1.str) {
        // short filename never needs encoding
        printf("Content-Disposition: attachment; filename=\"%s\"\n\n", attach->filename1.str);
    }
    else {
        // no filename is inline
        printf("Content-Disposition: inline\n\n");
    }

    (void)pst_attach_to_file_base64(pst, attach, stdout);
    printf("\n\n");
    DEBUG_RET();
}


int  header_match(char *header, char*field) {
    int n = strlen(field);
    if (strncasecmp(header, field, n) == 0) return 1;   // tag:{space}
    if ((field[n-1] == ' ') && (strncasecmp(header, field, n-1) == 0)) {
        char *crlftab = "\r\n\t";
        DEBUG_INFO(("Possible wrapped header = %s\n", header));
        if (strncasecmp(header+n-1, crlftab, 3) == 0) return 1; // tag:{cr}{lf}{tab}
    }
    return 0;
}

int  valid_headers(char *header)
{
    // headers are sometimes really bogus - they seem to be fragments of the
    // message body, so we only use them if they seem to be real rfc822 headers.
    // this list is composed of ones that we have seen in real pst files.
    // there are surely others. the problem is - given an arbitrary character
    // string, is it a valid (or even reasonable) set of rfc822 headers?
    if (header) {
        if (header_match(header, "Content-Type: "                 )) return 1;
        if (header_match(header, "Date: "                         )) return 1;
        if (header_match(header, "From: "                         )) return 1;
        if (header_match(header, "MIME-Version: "                 )) return 1;
        if (header_match(header, "Microsoft Mail Internet Headers")) return 1;
        if (header_match(header, "Received: "                     )) return 1;
        if (header_match(header, "Return-Path: "                  )) return 1;
        if (header_match(header, "Subject: "                      )) return 1;
        if (header_match(header, "To: "                           )) return 1;
        if (header_match(header, "X-ASG-Debug-ID: "               )) return 1;
        if (header_match(header, "X-Barracuda-URL: "              )) return 1;
        if (header_match(header, "X-x: "                          )) return 1;
        if (strlen(header) > 2) {
            DEBUG_INFO(("Ignore bogus headers = %s\n", header));
        }
        return 0;
    }
    else return 0;
}


void header_has_field(char *header, char *field, int *flag)
{
    DEBUG_ENT("header_has_field");
    if (my_stristr(header, field) || (strncasecmp(header, field+1, strlen(field)-1) == 0)) {
        DEBUG_INFO(("header block has %s header\n", field+1));
        *flag = 1;
    }
    DEBUG_RET();
}


void header_get_subfield(char *field, const char *subfield, char *body_subfield, size_t size_subfield)
{
    if (!field) return;
    DEBUG_ENT("header_get_subfield");
    char search[60];
    snprintf(search, sizeof(search), " %s=", subfield);
    field++;
    char *n = header_end_field(field);
    char *s = my_stristr(field, search);
    if (n && s && (s < n)) {
        char *e, *f, save;
        s += strlen(search);    // skip over subfield=
        if (*s == '"') {
            s++;
            e = strchr(s, '"');
        }
        else {
            e = strchr(s, ';');
            f = strchr(s, '\n');
            if (e && f && (f < e)) e = f;
        }
        if (!e || (e > n)) e = n;   // use the trailing lf as terminator if nothing better
        save = *e;
        *e = '\0';
        snprintf(body_subfield, size_subfield, "%s", s);  // copy the subfield to our buffer
        *e = save;
        DEBUG_INFO(("body %s %s from headers\n", subfield, body_subfield));
    }
    DEBUG_RET();
}

char* header_get_field(char *header, char *field)
{
    char *t = my_stristr(header, field);
    if (!t && (strncasecmp(header, field+1, strlen(field)-1) == 0)) t = header;
    return t;
}


// return pointer to \n at the end of this header field,
// or NULL if this field goes to the end of the string.
char *header_end_field(char *field)
{
    char *e = strchr(field+1, '\n');
    while (e && ((e[1] == ' ') || (e[1] == '\t'))) {
        e = strchr(e+1, '\n');
    }
    return e;
}


void header_strip_field(char *header, char *field)
{
    char *t;
    while ((t = header_get_field(header, field))) {
        char *e = header_end_field(t);
        if (e) {
            if (t == header) e++;   // if *t is not \n, we don't want to keep the \n at *e either.
            while (*e != '\0') {
                *t = *e;
                t++;
                e++;
            }
            *t = '\0';
        }
        else {
            // this was the last header field, truncate the headers
            *t = '\0';
        }
    }
}


int  test_base64(char *body, size_t len)
{
    int b64 = 0;
    uint8_t *b = (uint8_t *)body;
    DEBUG_ENT("test_base64");
    while (len--) {
        if ((*b < 32) && (*b != 9) && (*b != 10)) {
            DEBUG_INFO(("found base64 byte %d\n", (int)*b));
            b64 = 1;
            break;
        }
        b++;
    }
    DEBUG_RET();
    return b64;
}


void find_rfc822_headers(char** extra_mime_headers)
{
    DEBUG_ENT("find_rfc822_headers");
    char *headers = *extra_mime_headers;
    if (headers) {
        char *temp, *t;
        while ((temp = strstr(headers, "\n\n"))) {
            temp[1] = '\0';
            t = header_get_field(headers, "\nContent-Type:");
            if (t) {
                t++;
                DEBUG_INFO(("found content type header\n"));
                char *n = strchr(t, '\n');
                char *s = strstr(t, ": ");
                char *e = strchr(t, ';');
                if (!e || (e > n)) e = n;
                if (s && (s < e)) {
                    s += 2;
                    if (!strncasecmp(s, RFC822, e-s)) {
                        headers = temp+2;   // found rfc822 header
                        DEBUG_INFO(("found 822 headers\n%s\n", headers));
                        break;
                    }
                }
            }
            //DEBUG_INFO(("skipping to next block after\n%s\n", headers));
            headers = temp+2;   // skip to next chunk of headers
        }
        *extra_mime_headers = headers;
    }
    DEBUG_RET();
}


void write_body_part(pst_string *body, char *mime, char *charset, char *boundary, pst_file* pst)
{
    DEBUG_ENT("write_body_part");
    removeCR(body->str);
    size_t body_len = strlen(body->str);

    if (body->is_utf8) {
            charset = "utf-8";
    }
    int base64 = test_base64(body->str, body_len);
    printf("\n--%s\n", boundary);
    printf("Content-Type: %s; charset=\"%s\"\n", mime, charset);
    if (base64) printf("Content-Transfer-Encoding: base64\n");
    printf("\n");
    // Any body that uses an encoding with NULLs, e.g. UTF16, will be base64-encoded here.
    if (base64) {
        char *enc = pst_base64_encode(body->str, body_len);
        if (enc) {
            fputs(enc, stdout);
            printf("\n");
            free(enc);
        }
    }
    else {
    	fputs(body->str, stdout);
    }
    DEBUG_RET();
}


void write_schedule_part_data(pst_item* item, const char* sender, const char* method)
{
    printf("BEGIN:VCALENDAR\n");
    printf("PRODID:LibPST\n");
    if (method) printf("METHOD:%s\n", method);
    printf("BEGIN:VEVENT\n");
    if (sender) {
        if (item->email->outlook_sender_name.str) {
            printf("ORGANIZER;CN=\"%s\":MAILTO:%s\n", item->email->outlook_sender_name.str, sender);
        } else {
            printf("ORGANIZER;CN=\"\":MAILTO:%s\n", sender);
        }
    }
    write_appointment(item);
    printf("END:VCALENDAR\n");
}


void write_schedule_part(pst_item* item, const char* sender, const char* boundary)
{
    const char* method  = "REQUEST";
    const char* charset = "utf-8";
    char fname[30];
    if (!item->appointment) return;

    // inline appointment request
    printf("\n--%s\n", boundary);
    printf("Content-Type: %s; method=\"%s\"; charset=\"%s\"\n\n", "text/calendar", method, charset);
    write_schedule_part_data(item, sender, method);
    printf("\n");

    // attachment appointment request
    snprintf(fname, sizeof(fname), "i%i.ics", rand());
    printf("\n--%s\n", boundary);
    printf("Content-Type: %s; charset=\"%s\"; name=\"%s\"\n", "text/calendar", "utf-8", fname);
    printf("Content-Disposition: attachment; filename=\"%s\"\n\n", fname);
    write_schedule_part_data(item, sender, method);
    printf("\n");
}


void write_normal_email(pst_item* item, pst_file* pst, int embedding, char** extra_mime_headers)
{
    char boundary[60];
    char altboundary[66];
    char *altboundaryp = NULL;
    char body_charset[30];
    char buffer_charset[30];
    char body_report[60];
    char sender[60];
    int  sender_known = 0;
    char *temp = NULL;
    time_t em_time;
    char *c_time;
    char *headers = NULL;
    int has_from, has_subject, has_to, has_cc, has_date, has_msgid;
    has_from = has_subject = has_to = has_cc = has_date = has_msgid = 0;
    DEBUG_ENT("write_normal_email");

    pst_convert_utf8_null(item, &item->email->header);
    headers = valid_headers(item->email->header.str) ? item->email->header.str :
              valid_headers(*extra_mime_headers)     ? *extra_mime_headers     :
              NULL;

    // setup default body character set and report type
    strncpy(body_charset, pst_default_charset(item, sizeof(buffer_charset), buffer_charset), sizeof(body_charset));
    body_charset[sizeof(body_charset)-1] = '\0';
    strncpy(body_report, "delivery-status", sizeof(body_report));
    body_report[sizeof(body_report)-1] = '\0';

    // setup default sender
    pst_convert_utf8(item, &item->email->sender_address);
    if (item->email->sender_address.str && strchr(item->email->sender_address.str, '@')) {
        temp = item->email->sender_address.str;
        sender_known = 1;
    }
    else {
        temp = "MAILER-DAEMON";
    }
    strncpy(sender, temp, sizeof(sender));
    sender[sizeof(sender)-1] = '\0';

    // convert the sent date if it exists, or set it to a fixed date
    if (item->email->sent_date) {
        em_time = pst_fileTimeToUnixTime(item->email->sent_date);
        c_time = ctime(&em_time);
        if (c_time)
            c_time[strlen(c_time)-1] = '\0'; //remove end \n
        else
            c_time = "Thu Jan 1 00:00:00 1970";
    } else
        c_time = "Thu Jan 1 00:00:00 1970";

    // create our MIME boundaries here.
    snprintf(boundary, sizeof(boundary), "--boundary-LibPST-iamunique-%i_-_-", rand());
    snprintf(altboundary, sizeof(altboundary), "alt-%s", boundary);

    // we will always look at the headers to discover some stuff
    if (headers ) {
        char *t;
        removeCR(headers);

        temp = strstr(headers, "\n\n");
        if (temp) {
            // cut off our real rfc822 headers here
            temp[1] = '\0';
            // pointer to all the embedded MIME headers.
            // we use these to find the actual rfc822 headers for embedded message/rfc822 mime parts
            // but only for the outermost message
            if (!*extra_mime_headers) *extra_mime_headers = temp+2;
            DEBUG_INFO(("Found extra mime headers\n%s\n", temp+2));
        }

        // Check if the headers have all the necessary fields
        header_has_field(headers, "\nFrom:",        &has_from);
        header_has_field(headers, "\nTo:",          &has_to);
        header_has_field(headers, "\nSubject:",     &has_subject);
        header_has_field(headers, "\nDate:",        &has_date);
        header_has_field(headers, "\nCC:",          &has_cc);
        header_has_field(headers, "\nMessage-Id:",  &has_msgid);

        // look for charset and report-type in Content-Type header
        t = header_get_field(headers, "\nContent-Type:");
        header_get_subfield(t, "charset", body_charset, sizeof(body_charset));
        header_get_subfield(t, "report-type", body_report, sizeof(body_report));

        // derive a proper sender email address
        if (!sender_known) {
            t = header_get_field(headers, "\nFrom:");
            if (t) {
                // assume address is on the first line, rather than on a continuation line
                t++;
                char *n = strchr(t, '\n');
                char *s = strchr(t, '<');
                char *e = strchr(t, '>');
                if (s && e && n && (s < e) && (e < n)) {
                char save = *e;
                *e = '\0';
                    snprintf(sender, sizeof(sender), "%s", s+1);
                *e = save;
                }
            }
        }

        // Strip out the mime headers and some others that we don't want to emit
        header_strip_field(headers, "\nMicrosoft Mail Internet Headers");
        header_strip_field(headers, "\nMIME-Version:");
        header_strip_field(headers, "\nContent-Type:");
        header_strip_field(headers, "\nContent-Transfer-Encoding:");
        header_strip_field(headers, "\nContent-class:");
        header_strip_field(headers, "\nX-MimeOLE:");
        header_strip_field(headers, "\nX-From_:");
    }

    DEBUG_INFO(("About to print Header\n"));

    if (item && item->subject.str) {
        pst_convert_utf8(item, &item->subject);
        DEBUG_INFO(("item->subject = %s\n", item->subject.str));
    }

    // print the supplied email headers
    if (headers) {
        int len = strlen(headers);
        if (len > 0) {
            printf("%s", headers);
            // make sure the headers end with a \n
            if (headers[len-1] != '\n') printf("\n");
            //char *h = headers;
            //while (*h) {
            //    char *e = strchr(h, '\n');
            //    int   d = 1;    // normally e points to trailing \n
            //    if (!e) {
            //        e = h + strlen(h);  // e points to trailing null
            //        d = 0;
            //    }
            //    // we could do rfc2047 encoding here if needed
            //    printf("%.*s\n", (int)(e-h), h);
            //    h = e + d;
            //}
        }
    }

    // record read status
    if ((item->flags & PST_FLAG_READ) == PST_FLAG_READ) {
        printf("Status: RO\n");
    }

    // create required header fields that are not already written

    if (!has_from) {
        if (item->email->outlook_sender_name.str){
            pst_rfc2047(item, &item->email->outlook_sender_name, 1);
            printf("From: %s <%s>\n", item->email->outlook_sender_name.str, sender);
        } else {
            printf("From: <%s>\n", sender);
        }
    }

    if (!has_subject) {
        if (item->subject.str) {
            pst_rfc2047(item, &item->subject, 0);
            printf("Subject: %s\n", item->subject.str);
        } else {
            printf("Subject: \n");
        }
    }

    if (!has_to && item->email->sentto_address.str) {
        pst_rfc2047(item, &item->email->sentto_address, 0);
        printf("To: %s\n", item->email->sentto_address.str);
    }

    if (!has_cc && item->email->cc_address.str) {
        pst_rfc2047(item, &item->email->cc_address, 0);
        printf("Cc: %s\n", item->email->cc_address.str);
    }

    if (!has_date && item->email->sent_date) {
        char c_time[C_TIME_SIZE];
        struct tm stm;
        gmtime_r(&em_time, &stm);
        strftime(c_time, C_TIME_SIZE, "%a, %d %b %Y %H:%M:%S %z", &stm);
        printf("Date: %s\n", c_time);
    }

    if (!has_msgid && item->email->messageid.str) {
        pst_convert_utf8(item, &item->email->messageid);
        printf("Message-Id: %s\n", item->email->messageid.str);
    }

    // add forensic headers to capture some .pst stuff that is not really
    // needed or used by mail clients
    pst_convert_utf8_null(item, &item->email->sender_address);
    if (item->email->sender_address.str && !strchr(item->email->sender_address.str, '@')
                                        && strcmp(item->email->sender_address.str, ".")
                                        && (strlen(item->email->sender_address.str) > 0)) {
        printf("X-libpst-forensic-sender: %s\n", item->email->sender_address.str);
    }

    if (item->email->bcc_address.str) {
        pst_convert_utf8(item, &item->email->bcc_address);
        printf("X-libpst-forensic-bcc: %s\n", item->email->bcc_address.str);
    }

    // add our own mime headers
    printf("MIME-Version: 1.0\n");
    if (item->type == PST_TYPE_REPORT) {
        // multipart/report for DSN/MDN reports
        printf("Content-Type: multipart/report; report-type=%s;\n\tboundary=\"%s\"\n", body_report, boundary);
    }
    else {
        printf("Content-Type: multipart/mixed;\n\tboundary=\"%s\"\n", boundary);
    }
    printf("\n");    // end of headers, start of body

    // now dump the body parts
    if ((item->type == PST_TYPE_REPORT) && (item->email->report_text.str)) {
        write_body_part(&item->email->report_text, "text/plain", body_charset, boundary, pst);
        printf("\n");
    }

    if (item->body.str && item->email->htmlbody.str) {
        // start the nested alternative part
        printf("\n--%s\n", boundary);
        printf("Content-Type: multipart/alternative;\n\tboundary=\"%s\"\n", altboundary);
        altboundaryp = altboundary;
    }
    else {
        altboundaryp = boundary;
    }

    if (item->body.str) {
        write_body_part(&item->body, "text/plain", body_charset, altboundaryp, pst);
    }

    if (item->email->htmlbody.str) {
        write_body_part(&item->email->htmlbody, "text/html", body_charset, altboundaryp, pst);
    }

    if (item->body.str && item->email->htmlbody.str) {
        // end the nested alternative part
        printf("\n--%s--\n", altboundary);
    }

    if (item->email->rtf_compressed.data) {
        pst_item_attach* attach = (pst_item_attach*)malloc_or_die(sizeof(pst_item_attach));
        DEBUG_INFO(("Adding RTF body as attachment\n"));
        memset(attach, 0, sizeof(pst_item_attach));
        attach->next = item->attach;
        item->attach = attach;
        attach->data.data         = pst_lzfu_decompress(item->email->rtf_compressed.data, item->email->rtf_compressed.size, &attach->data.size);
        attach->filename2.str     = strdup(RTF_ATTACH_NAME);
        attach->filename2.is_utf8 = 1;
        attach->mimetype.str      = strdup(RTF_ATTACH_TYPE);
        attach->mimetype.is_utf8  = 1;
    }

    if (item->email->encrypted_body.data) {
        pst_item_attach* attach = (pst_item_attach*)malloc_or_die(sizeof(pst_item_attach));
        DEBUG_INFO(("Adding encrypted text body as attachment\n"));
        memset(attach, 0, sizeof(pst_item_attach));
        attach->next = item->attach;
        item->attach = attach;
        attach->data.data = item->email->encrypted_body.data;
        attach->data.size = item->email->encrypted_body.size;
        item->email->encrypted_body.data = NULL;
    }

    if (item->email->encrypted_htmlbody.data) {
        pst_item_attach* attach = (pst_item_attach*)malloc_or_die(sizeof(pst_item_attach));
        DEBUG_INFO(("Adding encrypted HTML body as attachment\n"));
        memset(attach, 0, sizeof(pst_item_attach));
        attach->next = item->attach;
        item->attach = attach;
        attach->data.data = item->email->encrypted_htmlbody.data;
        attach->data.size = item->email->encrypted_htmlbody.size;
        item->email->encrypted_htmlbody.data = NULL;
    }

    if (item->type == PST_TYPE_SCHEDULE) {
        write_schedule_part(item, sender, boundary);
    }

    // other attachments
    {
        pst_item_attach* attach;
        int attach_num = 0;
        for (attach = item->attach; attach; attach = attach->next) {
            pst_convert_utf8_null(item, &attach->filename1);
            pst_convert_utf8_null(item, &attach->filename2);
            pst_convert_utf8_null(item, &attach->mimetype);
            DEBUG_INFO(("Attempting Attachment encoding\n"));
            if (attach->method == PST_ATTACH_EMBEDDED) {
                DEBUG_INFO(("have an embedded rfc822 message attachment\n"));
                if (attach->mimetype.str) {
                    DEBUG_INFO(("which already has a mime-type of %s\n", attach->mimetype.str));
                    free(attach->mimetype.str);
                }
                attach->mimetype.str = strdup(RFC822);
                attach->mimetype.is_utf8 = 1;
                find_rfc822_headers(extra_mime_headers);
                write_embedded_message(attach, boundary, pst, extra_mime_headers);
            }
            else if (attach->data.data || attach->i_id) {
		write_inline_attachment(attach, boundary, pst);
            }
        }
    }

    printf("\n--%s--\n\n", boundary);
    DEBUG_RET();
}


void write_vcard(pst_item* item, pst_item_contact* contact, char comment[])
{
    char*  result = NULL;
    size_t resultlen = 0;
    char   time_buffer[30];
    // We can only call rfc escape once per printf, since the second call
    // may free the buffer returned by the first call.
    // I had tried to place those into a single printf - Carl.

    DEBUG_ENT("write_vcard");

    // make everything utf8
    pst_convert_utf8_null(item, &contact->fullname);
    pst_convert_utf8_null(item, &contact->surname);
    pst_convert_utf8_null(item, &contact->first_name);
    pst_convert_utf8_null(item, &contact->middle_name);
    pst_convert_utf8_null(item, &contact->display_name_prefix);
    pst_convert_utf8_null(item, &contact->suffix);
    pst_convert_utf8_null(item, &contact->nickname);
    pst_convert_utf8_null(item, &contact->address1);
    pst_convert_utf8_null(item, &contact->address2);
    pst_convert_utf8_null(item, &contact->address3);
    pst_convert_utf8_null(item, &contact->home_po_box);
    pst_convert_utf8_null(item, &contact->home_street);
    pst_convert_utf8_null(item, &contact->home_city);
    pst_convert_utf8_null(item, &contact->home_state);
    pst_convert_utf8_null(item, &contact->home_postal_code);
    pst_convert_utf8_null(item, &contact->home_country);
    pst_convert_utf8_null(item, &contact->home_address);
    pst_convert_utf8_null(item, &contact->business_po_box);
    pst_convert_utf8_null(item, &contact->business_street);
    pst_convert_utf8_null(item, &contact->business_city);
    pst_convert_utf8_null(item, &contact->business_state);
    pst_convert_utf8_null(item, &contact->business_postal_code);
    pst_convert_utf8_null(item, &contact->business_country);
    pst_convert_utf8_null(item, &contact->business_address);
    pst_convert_utf8_null(item, &contact->other_po_box);
    pst_convert_utf8_null(item, &contact->other_street);
    pst_convert_utf8_null(item, &contact->other_city);
    pst_convert_utf8_null(item, &contact->other_state);
    pst_convert_utf8_null(item, &contact->other_postal_code);
    pst_convert_utf8_null(item, &contact->other_country);
    pst_convert_utf8_null(item, &contact->other_address);
    pst_convert_utf8_null(item, &contact->business_fax);
    pst_convert_utf8_null(item, &contact->business_phone);
    pst_convert_utf8_null(item, &contact->business_phone2);
    pst_convert_utf8_null(item, &contact->car_phone);
    pst_convert_utf8_null(item, &contact->home_fax);
    pst_convert_utf8_null(item, &contact->home_phone);
    pst_convert_utf8_null(item, &contact->home_phone2);
    pst_convert_utf8_null(item, &contact->isdn_phone);
    pst_convert_utf8_null(item, &contact->mobile_phone);
    pst_convert_utf8_null(item, &contact->other_phone);
    pst_convert_utf8_null(item, &contact->pager_phone);
    pst_convert_utf8_null(item, &contact->primary_fax);
    pst_convert_utf8_null(item, &contact->primary_phone);
    pst_convert_utf8_null(item, &contact->radio_phone);
    pst_convert_utf8_null(item, &contact->telex);
    pst_convert_utf8_null(item, &contact->job_title);
    pst_convert_utf8_null(item, &contact->profession);
    pst_convert_utf8_null(item, &contact->assistant_name);
    pst_convert_utf8_null(item, &contact->assistant_phone);
    pst_convert_utf8_null(item, &contact->company_name);
    pst_convert_utf8_null(item, &item->body);

    // the specification I am following is (hopefully) RFC2426 vCard Mime Directory Profile
    printf("BEGIN:VCARD\n");
    printf("FN:%s\n", pst_rfc2426_escape(contact->fullname.str, &result, &resultlen));

    //printf("N:%s;%s;%s;%s;%s\n",
    printf("N:%s;", (!contact->surname.str)             ? "" : pst_rfc2426_escape(contact->surname.str, &result, &resultlen));
    printf("%s;",   (!contact->first_name.str)          ? "" : pst_rfc2426_escape(contact->first_name.str, &result, &resultlen));
    printf("%s;",   (!contact->middle_name.str)         ? "" : pst_rfc2426_escape(contact->middle_name.str, &result, &resultlen));
    printf("%s;",   (!contact->display_name_prefix.str) ? "" : pst_rfc2426_escape(contact->display_name_prefix.str, &result, &resultlen));
    printf("%s\n",  (!contact->suffix.str)              ? "" : pst_rfc2426_escape(contact->suffix.str, &result, &resultlen));

    if (contact->nickname.str)
        printf("NICKNAME:%s\n", pst_rfc2426_escape(contact->nickname.str, &result, &resultlen));
    if (contact->address1.str)
        printf("EMAIL:%s\n", pst_rfc2426_escape(contact->address1.str, &result, &resultlen));
    if (contact->address2.str)
        printf("EMAIL:%s\n", pst_rfc2426_escape(contact->address2.str, &result, &resultlen));
    if (contact->address3.str)
        printf("EMAIL:%s\n", pst_rfc2426_escape(contact->address3.str, &result, &resultlen));
    if (contact->birthday)
        printf("BDAY:%s\n", pst_rfc2425_datetime_format(contact->birthday, sizeof(time_buffer), time_buffer));

    if (contact->home_address.str) {
        //printf("ADR;TYPE=home:%s;%s;%s;%s;%s;%s;%s\n",
        printf("ADR;TYPE=home:%s;",  (!contact->home_po_box.str)      ? "" : pst_rfc2426_escape(contact->home_po_box.str, &result, &resultlen));
        printf("%s;",                ""); // extended Address
        printf("%s;",                (!contact->home_street.str)      ? "" : pst_rfc2426_escape(contact->home_street.str, &result, &resultlen));
        printf("%s;",                (!contact->home_city.str)        ? "" : pst_rfc2426_escape(contact->home_city.str, &result, &resultlen));
        printf("%s;",                (!contact->home_state.str)       ? "" : pst_rfc2426_escape(contact->home_state.str, &result, &resultlen));
        printf("%s;",                (!contact->home_postal_code.str) ? "" : pst_rfc2426_escape(contact->home_postal_code.str, &result, &resultlen));
        printf("%s\n",               (!contact->home_country.str)     ? "" : pst_rfc2426_escape(contact->home_country.str, &result, &resultlen));
        printf("LABEL;TYPE=home:%s\n", pst_rfc2426_escape(contact->home_address.str, &result, &resultlen));
    }

    if (contact->business_address.str) {
        //printf("ADR;TYPE=work:%s;%s;%s;%s;%s;%s;%s\n",
        printf("ADR;TYPE=work:%s;",  (!contact->business_po_box.str)      ? "" : pst_rfc2426_escape(contact->business_po_box.str, &result, &resultlen));
        printf("%s;",                ""); // extended Address
        printf("%s;",                (!contact->business_street.str)      ? "" : pst_rfc2426_escape(contact->business_street.str, &result, &resultlen));
        printf("%s;",                (!contact->business_city.str)        ? "" : pst_rfc2426_escape(contact->business_city.str, &result, &resultlen));
        printf("%s;",                (!contact->business_state.str)       ? "" : pst_rfc2426_escape(contact->business_state.str, &result, &resultlen));
        printf("%s;",                (!contact->business_postal_code.str) ? "" : pst_rfc2426_escape(contact->business_postal_code.str, &result, &resultlen));
        printf("%s\n",               (!contact->business_country.str)     ? "" : pst_rfc2426_escape(contact->business_country.str, &result, &resultlen));
        printf("LABEL;TYPE=work:%s\n", pst_rfc2426_escape(contact->business_address.str, &result, &resultlen));
    }

    if (contact->other_address.str) {
        //printf("ADR;TYPE=postal:%s;%s;%s;%s;%s;%s;%s\n",
        printf("ADR;TYPE=postal:%s;",(!contact->other_po_box.str)       ? "" : pst_rfc2426_escape(contact->other_po_box.str, &result, &resultlen));
        printf("%s;",                ""); // extended Address
        printf("%s;",                (!contact->other_street.str)       ? "" : pst_rfc2426_escape(contact->other_street.str, &result, &resultlen));
        printf("%s;",                (!contact->other_city.str)         ? "" : pst_rfc2426_escape(contact->other_city.str, &result, &resultlen));
        printf("%s;",                (!contact->other_state.str)        ? "" : pst_rfc2426_escape(contact->other_state.str, &result, &resultlen));
        printf("%s;",                (!contact->other_postal_code.str)  ? "" : pst_rfc2426_escape(contact->other_postal_code.str, &result, &resultlen));
        printf("%s\n",               (!contact->other_country.str)      ? "" : pst_rfc2426_escape(contact->other_country.str, &result, &resultlen));
        printf("LABEL;TYPE=postal:%s\n", pst_rfc2426_escape(contact->other_address.str, &result, &resultlen));
    }

    if (contact->business_fax.str)      printf("TEL;TYPE=work,fax:%s\n",         pst_rfc2426_escape(contact->business_fax.str, &result, &resultlen));
    if (contact->business_phone.str)    printf("TEL;TYPE=work,voice:%s\n",       pst_rfc2426_escape(contact->business_phone.str, &result, &resultlen));
    if (contact->business_phone2.str)   printf("TEL;TYPE=work,voice:%s\n",       pst_rfc2426_escape(contact->business_phone2.str, &result, &resultlen));
    if (contact->car_phone.str)         printf("TEL;TYPE=car,voice:%s\n",        pst_rfc2426_escape(contact->car_phone.str, &result, &resultlen));
    if (contact->home_fax.str)          printf("TEL;TYPE=home,fax:%s\n",         pst_rfc2426_escape(contact->home_fax.str, &result, &resultlen));
    if (contact->home_phone.str)        printf("TEL;TYPE=home,voice:%s\n",       pst_rfc2426_escape(contact->home_phone.str, &result, &resultlen));
    if (contact->home_phone2.str)       printf("TEL;TYPE=home,voice:%s\n",       pst_rfc2426_escape(contact->home_phone2.str, &result, &resultlen));
    if (contact->isdn_phone.str)        printf("TEL;TYPE=isdn:%s\n",             pst_rfc2426_escape(contact->isdn_phone.str, &result, &resultlen));
    if (contact->mobile_phone.str)      printf("TEL;TYPE=cell,voice:%s\n",       pst_rfc2426_escape(contact->mobile_phone.str, &result, &resultlen));
    if (contact->other_phone.str)       printf("TEL;TYPE=msg:%s\n",              pst_rfc2426_escape(contact->other_phone.str, &result, &resultlen));
    if (contact->pager_phone.str)       printf("TEL;TYPE=pager:%s\n",            pst_rfc2426_escape(contact->pager_phone.str, &result, &resultlen));
    if (contact->primary_fax.str)       printf("TEL;TYPE=fax,pref:%s\n",         pst_rfc2426_escape(contact->primary_fax.str, &result, &resultlen));
    if (contact->primary_phone.str)     printf("TEL;TYPE=phone,pref:%s\n",       pst_rfc2426_escape(contact->primary_phone.str, &result, &resultlen));
    if (contact->radio_phone.str)       printf("TEL;TYPE=pcs:%s\n",              pst_rfc2426_escape(contact->radio_phone.str, &result, &resultlen));
    if (contact->telex.str)             printf("TEL;TYPE=bbs:%s\n",              pst_rfc2426_escape(contact->telex.str, &result, &resultlen));
    if (contact->job_title.str)         printf("TITLE:%s\n",                     pst_rfc2426_escape(contact->job_title.str, &result, &resultlen));
    if (contact->profession.str)        printf("ROLE:%s\n",                      pst_rfc2426_escape(contact->profession.str, &result, &resultlen));
    if (contact->assistant_name.str || contact->assistant_phone.str) {
        printf("AGENT:BEGIN:VCARD\n");
        if (contact->assistant_name.str)    printf("FN:%s\n",                    pst_rfc2426_escape(contact->assistant_name.str, &result, &resultlen));
        if (contact->assistant_phone.str)   printf("TEL:%s\n",                   pst_rfc2426_escape(contact->assistant_phone.str, &result, &resultlen));
    }
    if (contact->company_name.str)      printf("ORG:%s\n",                       pst_rfc2426_escape(contact->company_name.str, &result, &resultlen));
    if (comment)                        printf("NOTE:%s\n",                      pst_rfc2426_escape(comment, &result, &resultlen));
    if (item->body.str)                 printf("NOTE:%s\n",                      pst_rfc2426_escape(item->body.str, &result, &resultlen));

    write_extra_categories(item);

    printf("VERSION: 3.0\n");
    printf("END:VCARD\n\n");
    if (result) free(result);
    DEBUG_RET();
}


/**
 * write extra vcard or vcalendar categories from the extra keywords fields
 *
 * @param item     pst item containing the keywords
 * @return         true if we write a categories line
 */
int write_extra_categories(pst_item* item)
{
    char*  result = NULL;
    size_t resultlen = 0;
    pst_item_extra_field *ef = item->extra_fields;
    const char *fmt = "CATEGORIES:%s";
    int category_started = 0;
    while (ef) {
        if (strcmp(ef->field_name, "Keywords") == 0) {
            printf(fmt, pst_rfc2426_escape(ef->value, &result, &resultlen));
            fmt = ", %s";
            category_started = 1;
        }
        ef = ef->next;
    }
    if (category_started) printf("\n");
    if (result) free(result);
    return category_started;
}


void write_journal(pst_item* item)
{
    char*  result = NULL;
    size_t resultlen = 0;
    char   time_buffer[30];
    pst_item_journal* journal = item->journal;

    // make everything utf8
    pst_convert_utf8_null(item, &item->subject);
    pst_convert_utf8_null(item, &item->body);

    printf("BEGIN:VJOURNAL\n");
    printf("DTSTAMP:%s\n",                     pst_rfc2445_datetime_format_now(sizeof(time_buffer), time_buffer));
    if (item->create_date)
        printf("CREATED:%s\n",                 pst_rfc2445_datetime_format(item->create_date, sizeof(time_buffer), time_buffer));
    if (item->modify_date)
        printf("LAST-MOD:%s\n",                pst_rfc2445_datetime_format(item->modify_date, sizeof(time_buffer), time_buffer));
    if (item->subject.str)
        printf("SUMMARY:%s\n",                 pst_rfc2426_escape(item->subject.str, &result, &resultlen));
    if (item->body.str)
        printf("DESCRIPTION:%s\n",             pst_rfc2426_escape(item->body.str, &result, &resultlen));
    if (journal && journal->start)
        printf("DTSTART;VALUE=DATE-TIME:%s\n", pst_rfc2445_datetime_format(journal->start, sizeof(time_buffer), time_buffer));
    printf("END:VJOURNAL\n");
    if (result) free(result);
}


void write_appointment(pst_item* item)
{
    char*  result = NULL;
    size_t resultlen = 0;
    char   time_buffer[30];
    pst_item_appointment* appointment = item->appointment;

    // make everything utf8
    pst_convert_utf8_null(item, &item->subject);
    pst_convert_utf8_null(item, &item->body);
    pst_convert_utf8_null(item, &appointment->location);

    printf("UID:%#"PRIx64"\n", item->block_id);
    printf("DTSTAMP:%s\n",                     pst_rfc2445_datetime_format_now(sizeof(time_buffer), time_buffer));
    if (item->create_date)
        printf("CREATED:%s\n",                 pst_rfc2445_datetime_format(item->create_date, sizeof(time_buffer), time_buffer));
    if (item->modify_date)
        printf("LAST-MOD:%s\n",                pst_rfc2445_datetime_format(item->modify_date, sizeof(time_buffer), time_buffer));
    if (item->subject.str)
        printf("SUMMARY:%s\n",                 pst_rfc2426_escape(item->subject.str, &result, &resultlen));
    if (item->body.str)
        printf("DESCRIPTION:%s\n",             pst_rfc2426_escape(item->body.str, &result, &resultlen));
    if (appointment && appointment->start)
        printf("DTSTART;VALUE=DATE-TIME:%s\n", pst_rfc2445_datetime_format(appointment->start, sizeof(time_buffer), time_buffer));
    if (appointment && appointment->end)
        printf("DTEND;VALUE=DATE-TIME:%s\n",   pst_rfc2445_datetime_format(appointment->end, sizeof(time_buffer), time_buffer));
    if (appointment && appointment->location.str)
        printf("LOCATION:%s\n",                pst_rfc2426_escape(appointment->location.str, &result, &resultlen));
    if (appointment) {
        switch (appointment->showas) {
            case PST_FREEBUSY_TENTATIVE:
                printf("STATUS:TENTATIVE\n");
                break;
            case PST_FREEBUSY_FREE:
                // mark as transparent and as confirmed
                printf("TRANSP:TRANSPARENT\n");
            case PST_FREEBUSY_BUSY:
            case PST_FREEBUSY_OUT_OF_OFFICE:
                printf("STATUS:CONFIRMED\n");
                break;
        }
        if (appointment->is_recurring) {
            const char* rules[] = {"DAILY", "WEEKLY", "MONTHLY", "YEARLY"};
            const char* days[]  = {"SU", "MO", "TU", "WE", "TH", "FR", "SA"};
            pst_recurrence *rdata = pst_convert_recurrence(appointment);
            printf("RRULE:FREQ=%s", rules[rdata->type]);
            if (rdata->count)       printf(";COUNT=%u",      rdata->count);
            if ((rdata->interval != 1) &&
                (rdata->interval))  printf(";INTERVAL=%u",   rdata->interval);
            if (rdata->dayofmonth)  printf(";BYMONTHDAY=%d", rdata->dayofmonth);
            if (rdata->monthofyear) printf(";BYMONTH=%d",    rdata->monthofyear);
            if (rdata->position)    printf(";BYSETPOS=%d",   rdata->position);
            if (rdata->bydaymask) {
                char byday[40];
                int  empty = 1;
                int i=0;
                memset(byday, 0, sizeof(byday));
                for (i=0; i<6; i++) {
                    int bit = 1 << i;
                    if (bit & rdata->bydaymask) {
                        char temp[40];
                        snprintf(temp, sizeof(temp), "%s%s%s", byday, (empty) ? ";BYDAY=" : ";", days[i]);
                        strcpy(byday, temp);
                        empty = 0;
                    }
                }
                printf("%s", byday);
            }
            printf("\n");
            pst_free_recurrence(rdata);
        }
        switch (appointment->label) {
            case PST_APP_LABEL_NONE:
                if (!write_extra_categories(item)) printf("CATEGORIES:NONE\n");
                break;
            case PST_APP_LABEL_IMPORTANT:
                printf("CATEGORIES:IMPORTANT\n");
                break;
            case PST_APP_LABEL_BUSINESS:
                printf("CATEGORIES:BUSINESS\n");
                break;
            case PST_APP_LABEL_PERSONAL:
                printf("CATEGORIES:PERSONAL\n");
                break;
            case PST_APP_LABEL_VACATION:
                printf("CATEGORIES:VACATION\n");
                break;
            case PST_APP_LABEL_MUST_ATTEND:
                printf("CATEGORIES:MUST-ATTEND\n");
                break;
            case PST_APP_LABEL_TRAVEL_REQ:
                printf("CATEGORIES:TRAVEL-REQUIRED\n");
                break;
            case PST_APP_LABEL_NEEDS_PREP:
                printf("CATEGORIES:NEEDS-PREPARATION\n");
                break;
            case PST_APP_LABEL_BIRTHDAY:
                printf("CATEGORIES:BIRTHDAY\n");
                break;
            case PST_APP_LABEL_ANNIVERSARY:
                printf("CATEGORIES:ANNIVERSARY\n");
                break;
            case PST_APP_LABEL_PHONE_CALL:
                printf("CATEGORIES:PHONE-CALL\n");
                break;
        }
        // ignore bogus alarms
        if (appointment->alarm && (appointment->alarm_minutes >= 0) && (appointment->alarm_minutes < 1440)) {
            printf("BEGIN:VALARM\n");
            printf("TRIGGER:-PT%dM\n", appointment->alarm_minutes);
            printf("ACTION:DISPLAY\n");
            printf("DESCRIPTION:Reminder\n");
            printf("END:VALARM\n");
        }
    }
    printf("END:VEVENT\n");
    if (result) free(result);
}


int main(int argc, char* const* argv) {
    pst_item *item = NULL;
    pst_desc_tree *d_ptr;
    int c,x;
    char *temp = NULL;               //temporary char pointer
    prog_name = argv[0];

    time_t now = time(NULL);
    srand((unsigned)now);

    mime_boundary = argv[1];
    json_template = argv[2];

    pst_file pstfile;
    if (pst_open(&pstfile, "input.blob", NULL)) {
    	    die("error opening PST");
    }
    if (pst_load_index(&pstfile)) {
    	    die("error loading PST index");
    }
    pst_load_extended_attributes(&pstfile);

    d_ptr = pstfile.d_head; // first record is main record
    item  = pst_parse_item(&pstfile, d_ptr, NULL);
    if (!item || !item->message_store) {
        die("Could not get root record");
    }

    // default the file_as to the same as the main filename if it doesn't exist
    if (!item->file_as.str) {
        item->file_as.str = (char*)malloc_or_die(strlen("pst")+1);
        strcpy(item->file_as.str, "pst");
        item->file_as.is_utf8 = 1;
        DEBUG_INFO(("file_as was blank, so am using %s\n", item->file_as.str));
    }
    DEBUG_INFO(("Root Folder Name: %s\n", item->file_as.str));

    d_ptr = pst_getTopOfFolders(&pstfile, item);
    if (!d_ptr) {
        die("Top of folders record not found.");
    }

    Progress progress;
    progress.n_processed = 0;
    progress.n_total = (item->folder) ? item->folder->item_count : 0;

    process(&pstfile, item, d_ptr->child, 0, "", &progress);    // do the children of TOPF

    pst_freeItem(item);
    pst_close(&pstfile);

    return 0;
}

