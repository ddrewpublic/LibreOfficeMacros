#!/bin/bash
# Extract StarBasic source from LibreOffice/OpenOffice .xba XML files
# and write plain-text .bas files next to them.
#
# NOTE: This version works when code is inside <script:module>...</script:module> tags.
# If your modules use <source>...</source> instead,
# this will produce empty files.
for f in *.xba; do
    name="${f%.xba}.bas"
    # sed -n '/<source>/,/<\/source>/{//!p}' "$f" \
    awk 'BEGIN{inmod=0}/<script:module/{inmod=1;next}/<\/script:module/{inmod=0}inmod{print}' "$f" \
    # Decode common XML entities back to literal characters for Basic code
    # &apos; -> '   &quot; -> "   &lt; -> <   &gt; -> >   &amp; -> &
    | sed -e 's/&apos;/'"'"'/g' -e 's/&quot;/"/g' \
    -e 's/&lt;/</g' -e 's/&gt;/>/g' -e 's/&amp;/\&/g' \
    # 3) Write result to the .bas file (overwrites if exists)
    > "$name"
done
