var maxBottom = 0;

function __getElementMaxBottom(element)
{
	if (element.className != "section")
	{
		var elementBottom = element.offsetTop + element.offsetHeight;
		if (elementBottom > maxBottom)
		{
			maxBottom = elementBottom;
		}
	}

	if (element.children != null)
	{	
		var children = element.children;
		
		for (var childIndex = 0; childIndex < children.length; childIndex++)
		{
			var child = children(childIndex);
			
			__getElementMaxBottom(child);
		}
	}
}

function __calcSectionHeight(section, minTop)
{
	maxBottom = 0;

	if (document.layers == null)
	{
		__getElementMaxBottom(section);
		section.style.height = maxBottom;

		absoluteBottom = section.offsetTop + section.offsetHeight;
		if (absoluteBottom < minTop)
		{
			diff = minTop - absoluteBottom;
			newHeight = parseInt(section.style.height) + diff;
			section.style.height = newHeight;
		}
	}
	else
	{
		absoluteBottom = section.pageY + section.clip.height;
		if (absoluteBottom < minTop)
		{
			section.resizeBy(0, (minTop - absoluteBottom));
		}
	}
}

