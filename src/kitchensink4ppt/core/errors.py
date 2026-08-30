"""Exception types for kitchensink4ppt. Every tool call maps these to actionable messages."""


class PptMcpError(Exception):
    """Base class; message text is user-facing."""


class DocumentNotFound(PptMcpError):
    pass


class DocumentLocked(PptMcpError):
    """File is open in PowerPoint (or another process holds a lock)."""


class DocumentCorrupt(PptMcpError):
    """File is not a valid .pptx package."""


class DocumentProtected(PptMcpError):
    """File is encrypted/password-protected."""


class TargetNotFound(PptMcpError):
    """Slide, shape, placeholder, or anchor the tool was told to act on does not exist."""


class AmbiguousTarget(PptMcpError):
    """More than one match for the given anchor; caller must disambiguate."""


class UnsupportedStructure(PptMcpError):
    """Package topology or XML shape we refuse to guess about (conservative mode)."""


class ValidationFailed(PptMcpError):
    """Post-edit validation caught a problem; the original file was NOT modified."""


class PowerPointNotRunning(PptMcpError):
    """No attachable interactive PowerPoint instance (live tools need one)."""


class DocumentNotOpenInPowerPoint(PptMcpError):
    """Live tool targeted a presentation that is not open in the running PowerPoint."""


class ProtectedViewRefused(PptMcpError):
    """Presentation is in Protected View; the user must click Enable Editing."""


class PowerPointBusy(PptMcpError):
    """PowerPoint rejected the call (dialog, Backstage, or a running command)."""


class PowerPointBlocked(PptMcpError):
    """PowerPoint is not answering at all (long synchronous operation in progress)."""


class PowerPointDisconnected(PptMcpError):
    """PowerPoint or the presentation closed mid-call; the edit may be partially applied."""
