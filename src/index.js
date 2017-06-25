import { saveAs } from 'file-saver';
import JSZip from 'jszip';

if (!typeof window === 'object') {
	throw new Error('Must be run in browser context');
}

document.addEventListener('DOMContentLoaded', function() {
	const picker = document.getElementById('picker');
	const palettes = document.getElementById('palettes');
	const button = document.getElementById('download');
	const parser = new DOMParser();
	const serializer = new XMLSerializer();
	const colorTypes = {
		dk1: 'Dark Color 1',
		dk2: 'Dark Color 2',
		lt1: 'Light Color 1',
		lt2: 'Light Color 2',
		accent1: 'Accent Color 1',
		accent2: 'Accent Color 2',
		accent3: 'Accent Color 3',
		accent4: 'Accent Color 4',
		accent5: 'Accent Color 5',
		accent6: 'Accent Color 6',
		hlink: 'Hyperlink Color',
		folHlink: 'Followed Hyperlink Color',
	};

	// remove all child nodes of an element
	function empty(el) {
		while (el.firstChild) {
			el.removeChild(el.firstChild);
		}
	}

	// clear out the selected file
	function clearFile() {
		picker.value = null;
		empty(palettes);
		button.setAttribute('disabled', 'disabled');
	}

	// when a file is picked
	picker.addEventListener('change', function() {
		if (picker.files.length === 0) {
			clearFile();
			return;
		}

		const file = picker.files[0];
		button.onclick = null;
		empty(palettes);

		// load the selected file
		JSZip.loadAsync(file).then(function(zip) {
			// grab the theme file(s) that contains the colors
			const files = zip.file(/ppt\/theme\/theme\d+.xml/);

			if (files.length === 0) {
				alert('No theme files found within - may not be a valid .pptx file');
				clearFile();
			}

			// process each file
			files.forEach(function(file) {
				// read and parse the contents of the file
				file.async('string').then(function(content) {
					button.removeAttribute('disabled');
					file.xml = parser.parseFromString(content, 'text/xml');

					// process each clrScheme element
					Array.prototype.forEach.call(file.xml.getElementsByTagName('clrScheme'), function(scheme) {
						// create a header for the color scheme
						const header = document.createElement('h3');
						header.textContent = scheme.getAttribute('name');
						palettes.appendChild(header);

						// create a list for all the colors of the palette
						const colorList = document.createElement('ul');

						// create a color picker for each color
						Array.prototype.forEach.call(scheme.children, function(color) {
							const li = document.createElement('li');
							const input = document.createElement('input');
							const colorType = color.tagName.split(':')[1]; // take the tagname after namespace
							const colorEl = color.children[0];

							input.setAttribute('type', 'color');
							input.setAttribute('value', '#' + colorEl.getAttribute('val'));

							// update theme xml when color picker changes
							input.addEventListener('change', function() {
								colorEl.setAttribute('val', input.value.substr(1).toUpperCase());
							});

							li.appendChild(input);
							li.appendChild(document.createTextNode(' ' + (colorTypes[colorType] || colorType)));
							colorList.appendChild(li);
						});

						palettes.appendChild(colorList);
					});
				});
			});

			// update download button action in the context of current file
			button.onclick = function() {
				button.setAttribute('disabled', 'disabled');

				// update the files
				files.forEach(function(file) {
					zip.file(file.name, serializer.serializeToString(file.xml));
				});

				// download new file
				zip.generateAsync({type: 'blob'}).then(function(content) {
					saveAs(content, file.name);
					button.removeAttribute('disabled');
				});
			};

		}, function(e) {
			alert(`Error reading ${file.name} : ${e.message}`);
			clearFile();
		});

	});

});
