# automation
Automate your SEO tasks
Internal Link Builder for Word docs
A small Python tool for SEOs and content teams: give it an Excel map of word → URL and a Word document, and it will automatically hyperlink the first occurrence of each word — once only — directly inside the .docx.

Great for building internal links at scale without doing it by hand.

What it does
Reads an .xlsx where column A = the anchor word and column B = the URL (a header row like Keyword | URL is detected and skipped automatically).
Links the first match of each word, once. No word is linked twice.
Whole-word, case-insensitive matching by default (Shoes matches shoes but not shoescraper).
Skips text that is already a hyperlink — never double-links.
Preserves formatting (bold, italic, size, etc.) around the new link.
When two words could match the same spot, the longer one wins (e.g. data analysis is linked instead of just data).
Works across body paragraphs and tables.
Install
pip install -r requirements.txt
Usage
python internal_linker.py --excel map.xlsx --docx article.docx
This writes article_linked.docx next to your original.

Options
Flag	Description
-e, --excel	Path to the .xlsx map (required).
-d, --docx	Path to the Word document (required).
-o, --output	Custom output path. Default: <name>_linked.docx.
--match-case	Make matching case-sensitive.
--color HEX	Link colour, e.g. 0563C1 (default) or none to keep original.
--no-underline	Don't underline the links.
--dry-run	Report what would be linked without saving a file.
Example Excel
Keyword	URL
technical seo	https://example.com/technical-seo
knowledge graph	https://example.com/knowledge-graph
internal linking	/guides/internal-linking
URLs can be absolute (https://…) or relative (/guides/…).

Good to know / limitations
A keyword that is split across multiple runs in the Word XML (e.g. half of it is bold) is skipped for that run — rare in normal documents, but possible.
"First occurrence" is resolved body paragraphs first, then tables.
Always review the output. Automation gets you 95% of the way; a human should still sanity-check anchor relevance.
Licence
MIT — free to use, change and share.
