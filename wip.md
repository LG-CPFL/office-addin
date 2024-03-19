## wip
- figure out how to do this
- is just building my own hotdocs viable?
- would need to create something that can parse script built into a template.

## holdingbay

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