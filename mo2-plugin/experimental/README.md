# Experimental Hooks

These modules are intentionally isolated from the stable tool menu integration.

Current experiment:

- `toolbar.py`: tries to add a direct LL Integration button to the MO2 main toolbar at runtime.

This uses Qt widget discovery rather than an official MO2 API. If it fails, the normal
`Tools > LL Integration` menu still works.
