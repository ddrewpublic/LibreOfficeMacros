for f in *.xba; do
  name="${f%.xba}.bas"
  awk 'BEGIN{inmod=0}/<script:module/{inmod=1;next}/<\/script:module/{inmod=0}inmod{print}' "$f" \
  | sed -e 's/&apos;/'"'"'/g' -e 's/&quot;/"/g' \
        -e 's/&lt;/</g' -e 's/&gt;/>/g' -e 's/&amp;/\&/g' > "$name"
done
