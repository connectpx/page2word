# Page2Word development plan

## Completed

1. **Extract articles** — Build the Manifest V3 extension and extract readable
   page content into editable sections.
2. **Edit locally** — Provide a bundled editor, autosave, section visibility,
   image replacement, and an optional normal-webpage editor for writing
   assistants.
3. **Export Word documents** — Generate formatted DOCX files with article
   text, tables, figures, and captured images.
4. **Configure formatting and links** — Store formatting options locally;
   remove original site links by default with an option to keep them. Provide
   navigable figure and table references in the editor and DOCX.
5. **Build and share** — Build the extension into `dist/extension`; document
   sharing the built-in editor and the optional Node-served editor.
6. **Publish releases** — Automate checks, build, package staging, and GitHub
   Release ZIP publication for version tags.

## Next

1. **Prepare a release candidate** — Run the documented test, typecheck, lint,
   and build commands; verify the generated extension loads from `dist/extension`.
2. **Manual Chrome validation** — Test extraction, formatting settings,
   source-link behavior, internal figure/table navigation, DOCX opening, and
   optional writing-assistant mode with the Node server both running and
   stopped.
3. **Publish the first release** — Validate the release workflow and publish
   the current `0.1.0` package from a matching `v0.1.0` tag.
4. **Consider Chrome Web Store distribution** — Follow Chrome's current
   publishing requirements if automatic extension updates are desired.
