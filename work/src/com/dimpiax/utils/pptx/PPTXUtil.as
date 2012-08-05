package com.dimpiax.utils.pptx {
	/**
	 * @author Pilipenko Dima
	 */
	public class PPTXUtil {
		public static function getSlideBody() : XML {
			var xml : XML = 
			<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
				<p:cSld>
					<p:spTree>
						<p:nvGrpSpPr>
							<p:cNvPr id="1" name=""/>
							<p:cNvGrpSpPr/>
							<p:nvPr/>
						</p:nvGrpSpPr>
						<p:grpSpPr>
							<a:xfrm>
								<a:off x="0" y="0"/>
								<a:ext cx="0" cy="0"/>
								<a:chOff x="0" y="0"/>
								<a:chExt cx="0" cy="0"/>
							</a:xfrm>
						</p:grpSpPr>
						<p:pic>
							<p:nvPicPr>
								<p:cNvPr id="4" name="filename" descr="This is image"/>
									<p:cNvPicPr>
										<a:picLocks noChangeAspect="1"/>
									</p:cNvPicPr>
								<p:nvPr/>
							</p:nvPicPr>
						<p:blipFill>
							<a:blip r:embed="rId2">
							<a:extLst>
								<a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
									<a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>
								</a:ext>
							</a:extLst>
							</a:blip>
							<a:stretch>
								<a:fillRect/>
							</a:stretch>
						</p:blipFill>
						<p:spPr>
							<a:xfrm>
								<a:off x="1741571" y="1165726"/>
								<a:ext cx="4457700" cy="2641600"/>
							</a:xfrm>
							<a:prstGeom prst="rect">
								<a:avLst/>
							</a:prstGeom>
						</p:spPr>
						</p:pic>
					</p:spTree>
					<p:extLst>
						<p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}">
							<p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="4109922657"/>
						</p:ext>
					</p:extLst>
				</p:cSld>
				<p:clrMapOvr>
					<a:masterClrMapping/>
				</p:clrMapOvr>
			</p:sld>;
			
			return xml.copy();
		}

		public static function getSlideRelBody() : XML {
			var xml : XML =
			<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
				<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
				<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="SET NAME IN CODE"/>
			</Relationships>;
			
			return xml.copy();
		}
	}
}
