# What's new

- XLSX exports now use a monospaced font throughout. The `Stations` column is
  focused on `Nn` and `Ns`, placing them at opposite ends of a compact,
  fixed-width cell while preserving removed-station brackets.
- XLSX exports again include a merged, framed ISO `Week` column and a blank
  `Transfer done` column for manual tracking.
- Sessions can now be added manually directly in the TUI with `A` or `+`.
  Fill Type, Code, Start, Dur, Stations, Ops, and Corr; `Enter` saves and
  `Esc` cancels.
- Select a manual session and press `Del` to remove it from
  `manual_sessions.json`.
- Use `--import-local` to import local `.vex` and `.skd` session files from the
  current directory. Imported sessions are stored in `manual_sessions.json` and
  skipped when the session code is already known.
- XLSX exports keep session code links clickable while matching the row/operator
  text color.
- Quick month filters are available with `F1`-`F12` for January through
  December. The same month filter can also be typed as `start: jun` or
  `start: june`.
- Startup default filters can still be stored in `startup_defaults.json`.
- Press `D` in the TUI to save the current filter as the startup default.
- Use `--no-mirrors` to override a startup default that enables mirror checks.
- Online help is available from the TUI, the manual page is included, and `--help` shows command-line options.

Have fun! --jole

Press `q`, `Enter`, or `Esc` to close this screen.
