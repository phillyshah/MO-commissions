"""The parser: the only code that opens an attachment (INT-5).

`engine.py` is the pure analysis (rasterize, text layer, hidden-text flags, risky-type check).
`service.py` wraps it as the Unix-socket server that runs in the no-network container.
`client.py` is what the rest of the portal calls; in test mode it runs the engine in-process.
"""
