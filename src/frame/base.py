class BaseSlide:
    """
    A base class for creating slides in a PowerPoint presentation with consistent
    formatting.
    """

    def __init__(self, slide_title: str) -> None: 
        self.slide_title = slide_title

    def add_slide(self):
        raise NotImplementedError("Subclasses should implement this!")