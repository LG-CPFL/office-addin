# hub
a company office add-in.
for what purpose?
currently unsure.
## wip
- figure out how to do this

      - name: Upload index
        uses: actions/upload-artifact@v4
        with: 
          name: index-artifact
          path: index.html
          overwrite: true

      - name: Upload dist
        uses: actions/upload-artifact@v4
        with:
          name: dist-artifact
          path: dist/
          overwrite: true

      - name: Merge artifacts
        uses: actions/upload-artifact/merge@v4
        with:
          name: github-pages
          delete-merged: true