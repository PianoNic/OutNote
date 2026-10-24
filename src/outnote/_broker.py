"""The FOCI token broker now lives in `entrance` (reused across projects).
This module re-exports it so OutNote's internal imports and the public
`TokenError` keep working unchanged."""

from entrance import GRAPH_SCOPE, Broker, TokenError, read_secret, write_secret

__all__ = ["Broker", "TokenError", "read_secret", "write_secret", "GRAPH_SCOPE"]
