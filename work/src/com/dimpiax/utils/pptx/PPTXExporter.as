package com.dimpiax.utils.pptx {
	import by.blooddy.crypto.image.JPEGEncoder;
	import by.blooddy.crypto.image.PNG24Encoder;

	import deng.fzip.FZip;
	import deng.fzip.FZipEvent;
	import deng.fzip.FZipFile;

	import flash.display.BitmapData;
	import flash.events.Event;
	import flash.events.EventDispatcher;
	import flash.events.IOErrorEvent;
	import flash.utils.ByteArray;

	/**
	 * @author Pilipenko Dima
	 */
	public class PPTXExporter extends EventDispatcher {
		private const CONTENT_TYPES : String = "[Content_Types].xml";
		
		private const PRESENTATION_RELS : String = "ppt/_rels/presentation.xml.rels";
		private const PRESENTATION : String = "ppt/presentation.xml";
		
		private const APPLICATION : String = "docProps/app.xml";
		
		//private var openXMLNS : Namespace = new Namespace("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		
		private var _zip : FZip;
		
		private var _filesCollector : Object;
		private var _imageNum : uint;
		
		public function PPTXExporter() {
			// create fzip
			_zip = new FZip();
			_zip.addEventListener(IOErrorEvent.IO_ERROR, function(event : IOErrorEvent) : void {trace("IO Error:", event.text);}, false, 0, true);
		}
		
		public function loadBytes(bytes : ByteArray) : void {
			initObjects();
			addListeners();
			
			_zip.loadBytes(bytes);
		}

		private function initObjects() : void {
			_filesCollector = {};
			_imageNum = 0;
		}
		
		public function addImage(image : BitmapData) : void {
			//trace("*** add image");
			
			var ext : String = ".png";
			var imageName : String = "image"+(_imageNum+1)+ext;
			var slideFilename : String = "slide"+(_imageNum+1)+".xml";
			
			var p : Namespace = new Namespace("http://schemas.openxmlformats.org/presentationml/2006/main");
			var r : Namespace = new Namespace("http://schemas.openxmlformats.org/package/2006/relationships");
			var a : Namespace = new Namespace("http://schemas.openxmlformats.org/drawingml/2006/main");
			
			// *** create slide
			var slideXML : XML = getSlideXML(_imageNum); // get needed slide
			
				// set properties of image
				var picNode : XMLList = slideXML.p::cSld.p::spTree.p::pic;
				
				// set description of image
				//picNode.p::nvPicPr.p::cNvPr.@descr = "The most beautiful image!";
				
				var imagePropsNode : XMLList = picNode.p::spPr.a::xfrm;
					imagePropsNode.a::off.@x = 0, imagePropsNode.a::off.@y = 613833; // set zero coordinates
					imagePropsNode.a::ext.@cx = 9144000, imagePropsNode.a::ext.@cy = 5518418; // set size on all slide of image
					
			// get needed slide relation
			var slideRelXML : XML = getSlideXML(_imageNum, true);
				slideRelXML.r::Relationship.(@Id == "rId2").@Target = "../media/"+imageName;
				//slideRelXML.appendChild(<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target={"../media/"+imageName}/>);
			
			// skip for first slide
			if(_imageNum) {
				// create slide in presentation
				var contentTypesXML : XML = _filesCollector[CONTENT_TYPES];
					contentTypesXML.appendChild(<Override PartName={"/ppt/slides/"+slideFilename} ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>);
					
				var presentationRelXML : XML = _filesCollector[PRESENTATION_RELS];
					var sldRID : String = "rId"+(presentationRelXML.*.length()+1);
					presentationRelXML.appendChild(<Relationship Id={sldRID} Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target={"slides/"+slideFilename} />);
				
				var r2 : Namespace = new Namespace("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
				var presentationXML : XML = _filesCollector[PRESENTATION];
				
					var slideIDList : XMLList = presentationXML.p::sldIdLst;
					var prevSldIDItem : XML = slideIDList.*[slideIDList.*.length()-1];
						slideIDList.appendChild(<p:sldId id={uint(prevSldIDItem.@id)+1} r:id={sldRID} xmlns:p={p.uri} xmlns:r={r2.uri}/>);
			}
			
			// save image in media folder
			var imageBA : ByteArray;
			if(ext == ".jpg" || ext == ".jpeg") imageBA = JPEGEncoder.encode(image, 100);
			else if(ext == ".png") imageBA = PNG24Encoder.encode(image);
			
			addFile("ppt/media/"+imageName, imageBA);
			
			// add slides
			addFile("ppt/slides/"+slideFilename, getBAFromXML(slideXML));
			addFile("ppt/slides/_rels/"+slideFilename+".rels", getBAFromXML(slideRelXML));
			
			_imageNum++;
		}
		
		// add file to archive, but remove first if exists file with this filename
		private function addFile(filename : String, ba : ByteArray) : void {
			var file : FZipFile = _zip.getFileByName(filename);
			if(file) {
				var index : int = _zip.getFileCount();
				while(index--) {
					if(file == _zip.getFileAt(index)) {
						_zip.removeFileAt(index);
						break;
					}
				}
			}
			
			file = _zip.addFile(filename, ba, false);
			file.date = new Date(1979, 11, 31);
		}
			
		private function getBAFromXML(xml : XML) : ByteArray {
			var ba : ByteArray = new ByteArray();
				
				var outputString:String = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
					outputString += xml.toXMLString(); 
					//outputString = outputString.replace(/\n/g, File.lineEnding);
			
				ba.writeUTFBytes(outputString);
			return ba;
		}

		private function getSlideXML(num : uint, rel : Boolean = false) : XML {
			var xml : XML, str : String;
			
			if(!rel) {
				str = "ppt/slides/slide"+(num+1)+".xml";
				xml = PPTXUtil.getSlideBody();
			}
			else {
				str = "ppt/slides/_rels/slide"+(num+1)+".xml.rels";
				xml = PPTXUtil.getSlideRelBody();
			}
			
			return xml;
		}

		private function addListeners(value : Boolean = true) : void {
			if(value) {
				_zip.addEventListener(FZipEvent.FILE_LOADED, onUnzipFileListener);
				_zip.addEventListener(Event.COMPLETE, onCompleteUnzipListener);
			}
			else {
				_zip.removeEventListener(FZipEvent.FILE_LOADED, onUnzipFileListener);
				_zip.removeEventListener(Event.COMPLETE, onCompleteUnzipListener);
				
			}
		}

		private function onUnzipFileListener(event : FZipEvent) : void {
			var file : FZipFile = event.file;
			var filename : String = file.filename;
			var fileBytes : ByteArray = file.content;
			
			
			switch(filename) {
				case CONTENT_TYPES: // need to add node of slide(name/type)
				
				case PRESENTATION: // need to add node of slide with specific rId
				case PRESENTATION_RELS: // register slide rId in list with name
				
				case APPLICATION: // change slides count in Properties.Slides
				
				// template of slide
				//case "ppt/slides/slide1.xml":
				//case "ppt/slides/_rels/slide1.xml.rels":
				
				_filesCollector[filename] = XML(fileBytes.readUTFBytes(fileBytes.bytesAvailable));
			}
		}

		private function onCompleteUnzipListener(event : Event) : void {
			dispatchEvent(new Event(Event.COMPLETE));
		}
		
		private function clear() : void {
			_filesCollector = null;
			_imageNum = 0;
		}
		
		public function pack() : FZip {
			// change thumbnail
			var file : FZipFile = _zip.getFileByName("ppt/media/image1.png");
			addFile("docProps/thumbnail.jpeg", file.content);
			file.content.position = 0;
			
			// update slides count
			var extPropertiesNS : Namespace = new Namespace("http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
			var appXML : XML = _filesCollector[APPLICATION];
				appXML.extPropertiesNS::Slides = _imageNum;
				
			var contentTypesXML : XML = _filesCollector[CONTENT_TYPES];
				contentTypesXML.appendChild(<Default Extension="png" ContentType="image/png"/>);
			
			// update files
			_zip.getFileByName(CONTENT_TYPES).content = getBAFromXML(_filesCollector[CONTENT_TYPES]);
			_zip.getFileByName(PRESENTATION_RELS).content = getBAFromXML(_filesCollector[PRESENTATION_RELS]);
			_zip.getFileByName(PRESENTATION).content = getBAFromXML(_filesCollector[PRESENTATION]);
			_zip.getFileByName(APPLICATION).content = getBAFromXML(appXML);
			
			addListeners(false);
			clear();
			
			return _zip;
		}

		public function close() : void {
			_zip = null;
		}
	}
}
