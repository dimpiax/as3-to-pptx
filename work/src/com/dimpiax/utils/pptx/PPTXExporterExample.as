package com.dimpiax.utils.pptx {
	import deng.fzip.FZip;

	import flash.display.Sprite;
	import flash.events.Event;

	/**
	 * @author Pilipenko Dima
	 */
	public class PPTXExporterExample extends Sprite {
		[Embed(source="../../../../../bin/model.pptx", mimeType="application/octet-stream")] public static var PPTX_MODEL : Class;
		
		private var _pptxExporter : PPTXExporter;
		
		public function PPTXExporterExample() {
			var pptxExporter : PPTXExporter = new PPTXExporter();
			_pptxExporter = pptxExporter;
					
				pptxExporter.addEventListener(Event.COMPLETE, onCompleteMouldModelListener);
				pptxExporter.loadBytes(new PPTX_MODEL());
		}

		private function onCompleteMouldModelListener(event : Event) : void {
			// add some image
			_pptxExporter.addImage(SOME_BITMAP_DATA);
			
			// arhive of moulded model
			var pptxZip : FZip = _pptxExporter.pack();
			
			// clear and remove pptx exporter
			_pptxExporter.close();
			_pptxExporter = null;
		}
	}
}
