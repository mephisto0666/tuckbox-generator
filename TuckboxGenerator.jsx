var dpi = 300.0;

function Card( width, height, thickness ) {
    this._width = width;
    this._height = height;
    this._thickness = thickness;
};

Card.prototype = {
    width: function() {
		return this._width;
	},
    height: function() {
		return this._height;
	},
    thickness: function() {
		return this._thickness;
	}
};

var ffgMiniSleevedCard = new Card( 44, 67, 0.55 );
var ffgRegularSleevedCard = new Card( 59, 92, 0.55 );
var dragonShieldStandardSleevedCard = new Card( 67, 91, 0.58 );

function BoxText( text, height ) {
    this._text = text;
    this._height = height;
};
	
BoxText.prototype = {
    text: function() {
		return this._text;
	},
	height: function() {
		return this._height;
	}
};

function TuckBox ( cardAmount, boxTexts, imageFilePath, cardFilePaths, logoFilePath, infoTexts, card, fileName ) {
    this._cardAmount = cardAmount;
    this._boxTexts = boxTexts;
    this._imageFilePath = imageFilePath;
    this._cardFilePaths = cardFilePaths;
    this._logoFilePath = logoFilePath;
    this._infoTexts = infoTexts;
    this._card = card;
    this._fileName = fileName;
    
    this._width = new UnitValue( this._card.width() + 2, "mm" );
    this._height = new UnitValue( this._card.height() + 2, "mm" );
    this._depth = new UnitValue( Math.ceil( this._cardAmount * this._card.thickness() ), "mm" );
    this._lip = new UnitValue( Math.ceil( Math.min( 20, this._width * 0.8 ) ), "mm" );
    this._bleed = new UnitValue( 2, "mm" );
    this._totalHeight = new UnitValue( this._depth + this._depth + this._width + this._width + this._lip, "mm" );
    this._totalWidth = new UnitValue( this._height + this._depth + this._depth, "mm" );
};

TuckBox.prototype = {	
	cardAmount: function() { 
		return this._cardAmount; 
	},
	boxTexts: function() { 
		return this._boxTexts; 
	},
	imageFilePath: function() { 
		return this._imageFilePath; 
	},
	cardFilePaths: function() { 
		return this._cardFilePaths; 
	},
	logoFilePath: function() { 
		return this._logoFilePath; 
	},
	infoTexts: function() { 
		return this._infoTexts; 
	},
	card: function() { 
		return this._card; 
	},
	fileName: function() { 
		return this._fileName; 
	},
	width: function() {
		return this._width;
	},
	height: function() {
		return this._height;
	},
	depth: function() {
		return this._depth;
	},
	lip: function() {
		return this._lip;
	},
	bleed: function() {
		return this._bleed;
	},
	totalHeight: function() {
		return this._totalHeight;
	},
	totalWidth: function() {
		return this._totalWidth;
	}
};

const documentSize = {
	A4: 1,
	LETTER: 2
};

function Vector( x, y ) {
	this.x = x;
	this.y = y;
}
Vector.prototype = {
	length: function() {
		return Math.sqrt( ( this.x * this.x ) + ( this.y * this.y ) );
	}
};

function TuckboxDrawer( dpi, documentSize ) {
    this._dpi = dpi;
    this._documentSize = documentSize;
};
TuckboxDrawer.prototype = {
	drawTuckbox: function( tuckbox ) {
		this.setSettings();
		this.initialize( tuckbox );
		this.prepareReferenceImage( tuckbox );
         this.createDocument( tuckbox.fileName() );
		this.createImageLayers( tuckbox );
		this.closeImageDocRef();
		this.drawLogo( tuckbox );
		this.drawCards( tuckbox );
		this.drawText( tuckbox );
		this.drawSinglePageBox( tuckbox );
		this.drawPageInfo( tuckbox.infoTexts(), tuckbox.cardAmount() );
		this.cleanup();
		this.restoreSettings();
	},
	setSettings: function() {
		this._defaultRulerUnits = preferences.rulerUnits;
		preferences.rulerUnits = Units.MM;
		UnitValue.baseUnit = UnitValue( 1/this._dpi, "in" );
	},
	initialize: function( tuckbox ) {
		this._fileRef = File( tuckbox.imageFilePath() );
		this._imageDocRef = app.open( this._fileRef );
	},
	prepareReferenceImage: function( tuckbox ) {
		this._imageDocRef.resizeImage( null, null, this._dpi, ResampleMethod.NONE );
		
		var totalHeight = tuckbox.totalHeight();
		var totalWidth = tuckbox.totalWidth();
		var imageHeight = totalHeight - tuckbox.width() + tuckbox.bleed();
		var imageWidth = totalWidth + tuckbox.bleed() + tuckbox.bleed();
		
		var imageRefHeight = this._imageDocRef.height;
		var imageRefWidth = this._imageDocRef.width;
		var resizeHeight = imageHeight / imageRefHeight;
		var resizeWidth = imageWidth / imageRefWidth;
		if( resizeHeight >= resizeWidth )
		{
			imageRefHeight = imageRefHeight * resizeHeight;
			imageRefWidth = imageRefWidth * resizeHeight;
		}
		else
		{
			imageRefHeight = imageRefHeight * resizeWidth;
			imageRefWidth = imageRefWidth * resizeWidth;
		}

		this._imageDocRef.flatten();
		this._imageDocRef.resizeImage( imageRefWidth, imageRefHeight, null, ResampleMethod.BICUBIC );
	},
	createDocument: function( fileName ) {
		switch( this._documentSize ) {
			case documentSize.A4:
				this._documentRef = documents.add( 210, 297, this._dpi, fileName );
				break;
			case documentSize.LETTER:
				preferences.rulerUnits = Units.INCHES;
				this._documentRef = documents.add( 8, 11, this._dpi, fileName );
				preferences.rulerUnits = Units.MM;
				break;
			default:
				throw new Error( "Document size not supported" );
		}
	},
	createImageLayers: function( tuckbox ) {
		var target_image_width = tuckbox.totalWidth() + tuckbox.bleed() + tuckbox.bleed();
		var target_topimage_height = tuckbox.totalHeight() - tuckbox.width() + tuckbox.bleed();
		var target_bottomimage_height = tuckbox.width() + tuckbox.bleed();
			
		var x_left = new UnitValue( Math.ceil( ( this._documentRef.width.as("mm") - target_image_width.as("mm") ) / 2 ), "mm" );
		var x_right = x_left + target_image_width;
		var y_top = new UnitValue( Math.ceil( ( ( this._documentRef.height.as("mm") - tuckbox.totalHeight().as("mm") ) / 2 ) - tuckbox.bleed().as("mm") ), "mm" );
		var y_bottom = y_top + target_topimage_height;
		
		app.activeDocument = this._imageDocRef;
		
		var image_height = this._imageDocRef.height;
		var image_width = this._imageDocRef.width;
		
		var image_x_left = ( image_width - target_image_width ) / 2;
		if(  image_x_left.as("mm") < 0 ) image_x_left *= 0;
		
		var image_x_right = image_x_left + target_image_width;
		if(  image_x_right > image_width ) image_x_right = image_width;
		
		var image_y_bottom = image_height;
		var image_y_top = image_y_bottom - target_topimage_height;
		if( image_y_top.as("mm") < 0 ) image_y_top *= 0;
		
		// Art Layer 1
		selRegion = Array( [image_x_left.as("px"), image_y_top.as("px")], 
							[image_x_right.as("px"), image_y_top.as("px")], 
							[image_x_right.as("px"), image_y_bottom.as("px")], 
							[image_x_left.as("px"), image_y_bottom.as("px")] );
		this._imageDocRef.selection.select( selRegion );
		this._imageDocRef.selection.copy();
		
		app.activeDocument = this._documentRef;
		
		var box_art_layer1 = this._documentRef.artLayers.add();
		box_art_layer1.name = "Box Art Layer 1";
		
		var selRegion = Array( [x_left.as("px"), y_top.as("px")], 
								[x_right.as("px"), y_top.as("px")], 
								[x_right.as("px"), y_bottom.as("px")], 
								[x_left.as("px"), y_bottom.as("px")] );
		this._documentRef.selection.select( selRegion );
		this._documentRef.paste();
		this._documentRef.selection.deselect();
		
		// Clear unneccesary parts of image	
		var clear_x_left = x_left.as( "px" ) + tuckbox.bleed().as( "px" ) + tuckbox.bleed().as( "px" );
		var clear_x_delta = tuckbox.depth().as( "px" ) - tuckbox.bleed().as( "px" ) - tuckbox.bleed().as( "px" );
		var clear_y_bottom = y_bottom.as( "px" ) - tuckbox.bleed().as( "px" );
		var clear_y_delta = tuckbox.depth().as( "px" ) - tuckbox.bleed().as( "px" ) - tuckbox.bleed().as( "px" );
		
		if( tuckbox.depth().as( "px" ) > ( tuckbox.height().as( "px" ) / 2 ) )
		{
			clear_x_left = x_left.as( "px" );
			clear_x_delta = tuckbox.depth().as( "px" );
			clear_y_bottom = y_bottom.as( "px" ) - tuckbox.bleed().as( "px" );
			clear_y_delta = tuckbox.depth().as( "px" ) - ( tuckbox.height().as( "px" ) * 0.5 ) - ( tuckbox.bleed().as( "px" ) * 2 );
			
			selRegion = Array( [clear_x_left, clear_y_bottom], 
								[clear_x_left + clear_x_delta, clear_y_bottom], 
								[clear_x_left + clear_x_delta, clear_y_bottom - clear_y_delta], 
								[clear_x_left, clear_y_bottom - clear_y_delta] );
			this._documentRef.selection.select( selRegion );
			this._documentRef.selection.clear();
			this._documentRef.selection.deselect();
			
			clear_x_left = x_left.as( "px" ) + tuckbox.bleed().as( "px" ) + tuckbox.bleed().as( "px" );
			clear_x_delta = tuckbox.depth().as( "px" ) - tuckbox.bleed().as( "px" ) - tuckbox.bleed().as( "px" );
			clear_y_bottom = y_bottom.as( "px" ) - tuckbox.bleed().as( "px" ) - tuckbox.depth().as( "px" ) + ( tuckbox.height().as( "px" ) * 0.5 );
			clear_y_delta = ( tuckbox.height().as( "px" ) * 0.5 ) - tuckbox.bleed().as( "px" ) - tuckbox.bleed().as( "px" );
		}
		
		selRegion = Array( [clear_x_left, clear_y_bottom], 
							[clear_x_left + clear_x_delta, clear_y_bottom], 
							[clear_x_left + clear_x_delta, clear_y_bottom - clear_y_delta], 
							[clear_x_left, clear_y_bottom - clear_y_delta] );
		this._documentRef.selection.select( selRegion );
		this._documentRef.selection.clear();
		this._documentRef.selection.deselect();
		
		clear_y_bottom = y_bottom.as( "px" ) - tuckbox.depth().as( "px" ) - tuckbox.bleed().as( "px" );
		clear_y_delta = tuckbox.width().as( "px" ) - tuckbox.bleed().as( "px" ) - tuckbox.bleed().as( "px" );
		
		selRegion = Array( [clear_x_left, clear_y_bottom], 
							[clear_x_left + clear_x_delta, clear_y_bottom], 
							[clear_x_left + clear_x_delta, clear_y_bottom - clear_y_delta], 
							[clear_x_left, clear_y_bottom - clear_y_delta] );
		this._documentRef.selection.select( selRegion );
		this._documentRef.selection.clear();
		this._documentRef.selection.deselect();
		
		clear_x_left = x_right.as( "px" ) - tuckbox.depth().as( "px" );
		clear_x_delta = tuckbox.depth().as( "px" ) - tuckbox.bleed().as( "px" ) - tuckbox.bleed().as( "px" );
		
		selRegion = Array( [clear_x_left, clear_y_bottom], 
							[clear_x_left + clear_x_delta, clear_y_bottom], 
							[clear_x_left + clear_x_delta, clear_y_bottom - clear_y_delta], 
							[clear_x_left, clear_y_bottom - clear_y_delta] );
		this._documentRef.selection.select( selRegion );
		this._documentRef.selection.clear();
		this._documentRef.selection.deselect();
		
		clear_y_bottom = y_bottom.as( "px" ) - tuckbox.bleed().as( "px" );
		clear_y_delta = tuckbox.depth().as( "px" ) - tuckbox.bleed().as( "px" ) - tuckbox.bleed().as( "px" );
		
		if( tuckbox.depth().as( "px" ) > ( tuckbox.height().as( "px" ) / 2 ) )
		{
			clear_x_left = x_right.as( "px" ) - tuckbox.depth().as( "px" );
			clear_x_delta = tuckbox.depth().as( "px" );
			clear_y_bottom = y_bottom.as( "px" ) - tuckbox.bleed().as( "px" );
			clear_y_delta = tuckbox.depth().as( "px" ) - ( tuckbox.height().as( "px" ) * 0.5 ) - ( tuckbox.bleed().as( "px" ) * 2 );
			
			selRegion = Array( [clear_x_left, clear_y_bottom], 
								[clear_x_left + clear_x_delta, clear_y_bottom], 
								[clear_x_left + clear_x_delta, clear_y_bottom - clear_y_delta], 
								[clear_x_left, clear_y_bottom - clear_y_delta] );
			this._documentRef.selection.select( selRegion );
			this._documentRef.selection.clear();
			this._documentRef.selection.deselect();
			
			clear_x_left = x_right.as( "px" ) - tuckbox.depth().as( "px" );
			clear_x_delta = tuckbox.depth().as( "px" ) - tuckbox.bleed().as( "px" ) - tuckbox.bleed().as( "px" );
			clear_y_bottom = y_bottom.as( "px" ) - tuckbox.bleed().as( "px" ) - tuckbox.depth().as( "px" ) + ( tuckbox.height().as( "px" ) * 0.5 );
			clear_y_delta = ( tuckbox.height().as( "px" ) * 0.5 ) - tuckbox.bleed().as( "px" ) - tuckbox.bleed().as( "px" );
		}
		
		selRegion = Array( [clear_x_left, clear_y_bottom], 
							[clear_x_left + clear_x_delta, clear_y_bottom], 
							[clear_x_left + clear_x_delta, clear_y_bottom - clear_y_delta], 
							[clear_x_left, clear_y_bottom - clear_y_delta] );
		this._documentRef.selection.select( selRegion );
		this._documentRef.selection.clear();
		this._documentRef.selection.deselect();
		
		// Art Layer 2
		app.activeDocument = this._imageDocRef;
		this._imageDocRef.selection.selectAll();
		this._imageDocRef.selection.rotate( 180 );
		selRegion = Array( [image_x_left.as( "px" ), tuckbox.depth().as( "px" )], 
							[image_x_right.as( "px" ), tuckbox.depth().as( "px" )], 
							[image_x_right.as( "px" ), target_bottomimage_height.as( "px" ) + tuckbox.depth().as( "px" )], 
							[image_x_left.as( "px" ), target_bottomimage_height.as( "px" ) + tuckbox.depth().as( "px" )] );
		this._imageDocRef.selection.select( selRegion );
		this._imageDocRef.selection.copy();
		app.activeDocument = this._documentRef;
		
		y_top = y_bottom;
		y_bottom = y_top + target_bottomimage_height;
		
		var box_art_layer2 = this._documentRef.artLayers.add();
		box_art_layer2.name = "Box Art Layer 2";
		
		selRegion = Array( [x_left.as( "px" ), y_top.as( "px" )], 
							[x_right.as( "px" ), y_top.as( "px" )], 
							[x_right.as( "px" ), y_bottom.as( "px" )], 
							[x_left.as( "px" ), y_bottom.as( "px" )] );
		this._documentRef.selection.select( selRegion );
		this._documentRef.paste();
		this._documentRef.selection.deselect();
		
		box_art_layer1 = null;
		box_art_layer2 = null;
	},
	closeImageDocRef: function() {
		this._imageDocRef.close( SaveOptions.DONOTSAVECHANGES );
	},
	drawLogo: function( tuckbox ) {
		var fileRef = File( tuckbox.logoFilePath() );
		var imageDocRef = app.open( fileRef );
		imageDocRef.resizeImage( null, null, dpi, ResampleMethod.NONE );

		app.activeDocument = imageDocRef;
		
		// Resize image for big logo on front
		var image_height = tuckbox.width() / 4;
		var image_width = tuckbox.height() / 2;
		
		var imageRef_height = imageDocRef.height;
		var imageRef_width = imageDocRef.width;
		
		var resize_height = image_height / imageRef_height;
		var resize_width = image_width / imageRef_width;
		if( resize_height >= resize_width )
		{
			imageRef_height = imageRef_height * resize_height;
			imageRef_width = imageRef_width * resize_height;
		}
		else
		{
			imageRef_height = imageRef_height * resize_width;
			imageRef_width = imageRef_width * resize_width;
		}

		imageDocRef.resizeImage( imageRef_width, imageRef_height, null, ResampleMethod.BICUBIC );
		imageDocRef.selection.selectAll();
		imageDocRef.selection.rotate( 180 );
		imageDocRef.selection.copy();
		
		// Create layer and paste big logo on front
		app.activeDocument = this._documentRef;
		
		var logo_art_layer1 = this._documentRef.artLayers.add();
		logo_art_layer1.name = "Logo Art Layer";
		
		var x_left = Math.ceil( ( this._documentRef.width.as( "px" ) - tuckbox.totalWidth().as( "px" ) ) / 2 );
		var x_right = x_left + tuckbox.totalWidth().as( "px" );
		var y_top = Math.ceil( ( this._documentRef.height.as( "px" ) - tuckbox.totalHeight().as( "px" ) ) / 2 );
		var y_bottom = y_top + tuckbox.totalHeight().as( "px" );
		
		var x1 = ( x_left + x_right ) / 2 - ( imageRef_width.as( "px" ) / 2 );
		var x2 = x1 + imageRef_width.as( "px" );
		var y1 = y_bottom - ( tuckbox.width().as("px") / 4 );
		var y2 = y1 - imageRef_height.as( "px" );
		
		var selRegion = Array( [x1, y1], [x2, y1], [x2, y2], [x1, y2] );
		this._documentRef.selection.select( selRegion );
		this._documentRef.paste();
		this._documentRef.selection.deselect();
		
		// Small logos on top and bottom
		app.activeDocument = imageDocRef;
		
		image_height = tuckbox.depth() / 4;
		image_width = tuckbox.height() / 4;
		
		resize_height = image_height / imageRef_height;
		resize_width = image_width / imageRef_width;
		if( resize_height >= resize_width )
		{
			imageRef_height = imageRef_height * resize_height;
			imageRef_width = imageRef_width * resize_height;
		}
		else
		{
			imageRef_height = imageRef_height * resize_width;
			imageRef_width = imageRef_width * resize_width;
		}

		imageDocRef.resizeImage( imageRef_width, imageRef_height, null, ResampleMethod.BICUBIC );
		imageDocRef.selection.selectAll();
		imageDocRef.selection.copy();
		
		app.activeDocument = this._documentRef;

		x1 = ( x_left + x_right ) / 2 + ( tuckbox.height().as("px") / 4 ) - ( imageRef_width.as("px") / 2 );
		x2 = x1 + imageRef_width.as("px");
		y1 = y_bottom - tuckbox.width().as("px") - tuckbox.depth().as("px") - tuckbox.width().as("px") - ( tuckbox.depth().as("px") / 2 )  + ( imageRef_height.as("px") / 2 );
		y2 = y1 - imageRef_height.as("px");
		
		selRegion = Array( [x1, y1], [x2, y1], [x2, y2], [x1, y2] );
		this._documentRef.selection.select( selRegion );
		this._documentRef.paste();
		this._documentRef.selection.deselect();
		
		app.activeDocument = imageDocRef;
		
		imageDocRef.selection.selectAll();
		imageDocRef.selection.rotate( 180 );
		imageDocRef.selection.copy();
		
		app.activeDocument = this._documentRef;
		
		x1 = ( x_left + x_right ) / 2 - ( tuckbox.height().as("px") / 4 ) - ( imageRef_width.as("px") / 2 );
		x2 = x1 + imageRef_width.as("px");
		y1 = y_bottom - tuckbox.width().as("px") - ( tuckbox.depth().as("px") / 2 ) + ( imageRef_height.as("px") / 2 );
		y2 = y1 - imageRef_height.as("px");
		
		selRegion = Array( [x1, y1], [x2, y1], [x2, y2], [x1, y2] );
		this._documentRef.selection.select( selRegion );
		this._documentRef.paste();
		this._documentRef.selection.deselect();
		
		imageDocRef.close( SaveOptions.DONOTSAVECHANGES );
		
		// Release references
		logo_art_layer1 = null;
		fileRef = null;
		imageDocRef = null;
	},
	drawCards: function( tuckbox ) {
		var cardFileCount = tuckbox.cardFilePaths().length;
		if( cardFileCount > 0 )
		{
			var x_distance = UnitValue( 10, "mm" );
			var y_distance = UnitValue( 2, "mm" );
			var angle_diff = 7;
			
			if( cardFileCount > 4 )
			{
				var factor = 4 / cardFileCount;
				x_distance *= factor;
				y_distance *= factor;
				angle_diff *= factor;
			}
			
			var x_left = Math.ceil( ( this._documentRef.width.as( "px" ) - tuckbox.totalWidth().as( "px" ) ) / 2 );
			var x_right = x_left + tuckbox.totalWidth().as( "px" );
			var y_top = Math.ceil( ( this._documentRef.height.as( "px" ) - tuckbox.totalHeight().as( "px" ) ) / 2 );
			var y_bottom = y_top + tuckbox.totalHeight().as( "px" );
			
			var angle = -5;
			
			var x1 = ( x_left + x_right ) / 2;
			var x2 = x1;
			var y1 = y_bottom - tuckbox.width().as( "px" ) - tuckbox.depth().as( "px" ) - ( tuckbox.width().as( "px" ) / 2 );
			var y2 = y1;
			
			var x_offset = 0;
			var y_offset = 0;
			
			var i = 0;
			
			for( i = 0; i < cardFileCount; i++ )
			{
				var imageFilePath = tuckbox.cardFilePaths()[i];
				var fileRef = File( imageFilePath );
				var imageDocRef = app.open( fileRef );
				imageDocRef.resizeImage( null, null, dpi, ResampleMethod.NONE );

				app.activeDocument = imageDocRef;
				
				// Resize image for big logo on front
				var image_height = tuckbox.width() * 0.5;
				
				var imageRef_height = imageDocRef.height;
				var imageRef_width = imageDocRef.width;
				
				var resize_height = image_height / imageRef_height;
				imageRef_height = imageRef_height * resize_height;
				imageRef_width = imageRef_width * resize_height;

				imageDocRef.resizeImage( imageRef_width, imageRef_height, null, ResampleMethod.BICUBIC );
				imageDocRef.selection.selectAll();
				imageDocRef.selection.copy();
				
				// Create layer and paste big logo on front
				app.activeDocument = this._documentRef;
				
				var logo_art_layer1 = this._documentRef.artLayers.add();
				logo_art_layer1.name = "Card " + i + " Art Layer";
				
				var middle = ( cardFileCount - 1 ) / 2;
				var temp = i - middle;
				
				x_offset = temp * x_distance.as( "px" );
				y_offset = Math.abs( temp * y_distance.as( "px" ) );
				
				x1 = ( x_left + x_right ) / 2 - ( imageRef_width.as( "px" ) / 2 ) + ( tuckbox.height().as( "px" ) * 0.03 );
				x2 = x1 + imageRef_width.as( "px" );
				y1 = y_bottom - tuckbox.width().as( "px" ) - tuckbox.depth().as( "px" ) - ( tuckbox.width().as( "px" ) * 0.6 ) + ( imageRef_height.as( "px" ) / 2 );
				y2 = y1 - imageRef_height.as( "px" );
				
				var selRegion = Array( [x1 + x_offset, y1 + y_offset], [x2 + x_offset, y1 + y_offset], [x2 + x_offset, y2 + y_offset], [x1 + x_offset, y2 + y_offset] );
				this._documentRef.selection.select( selRegion );
				this._documentRef.paste();
				this._documentRef.selection.selectAll();
				this._documentRef.selection.rotate( angle + ( angle_diff * temp ) );
				this._documentRef.selection.deselect();
				
				imageDocRef.close( SaveOptions.DONOTSAVECHANGES );
				
				// Release references
				logo_art_layer1 = null;
				fileRef = null;
				imageDocRef = null;
			}
		}
	},
	drawText: function( tuckbox ) {
		var textAmount = tuckbox.boxTexts().length;
		if( textAmount > 0 )
		{		
			var textColor = new SolidColor();
			textColor.rgb.red = 255;
			textColor.rgb.green = 255;
			textColor.rgb.blue = 255;
		
			var box_width_points = tuckbox.width();
			var box_height_points = tuckbox.height();
			var box_depth_points = tuckbox.depth();
			
			var x_left = Math.ceil((this._documentRef.width - tuckbox.totalWidth()) / 2 );
			var x_right = x_left + tuckbox.totalWidth();
			var y_top = Math.ceil((this._documentRef.height - tuckbox.totalHeight()) / 2 );
			var y_bottom = y_top + tuckbox.totalHeight();
			
			var x1 = ( x_left + x_right ) / 2;
			var x2 = x1;
			var y1 = y_bottom - box_width_points - box_depth_points - ( box_width_points / 2 );
			var y2 = y1;
			
			var offset = 0;
			var linespace = 2;
			var font = "AdobeDevanagari-Regular";
			
			var total_height = linespace * ( textAmount - 1 );
			for( i = textAmount - 1; i >= 0; i-- )
			{
				total_height += this.convertToMillimeter( tuckbox.boxTexts()[i].height() );
				if( i > 1 ) total_height += linespace;
			}
			var current_height = 0;
			
			var r = 255;
			var g = 255;
			var b = 255;
			var blendingMode = "Nrml";
			var	opacity = 75; 
			var spread = 0;
			var size = 5;
			
			var i = 0;
			for( i = textAmount - 1; i >= 0; i-- )
			{
				// FRONT
				var text_art_layer = this._documentRef.artLayers.add();
				text_art_layer.name = "Text " + i + " - Front Layer";
				text_art_layer.kind = LayerKind.TEXT;
				
				var txtRef = text_art_layer.textItem;
				txtRef.font = font;
				txtRef.contents = tuckbox.boxTexts()[i].text();
				txtRef.size = tuckbox.boxTexts()[i].height();
				txtRef.color = textColor;
				txtRef.kind = TextType.PARAGRAPHTEXT;
				
				text_art_layer.rotate( 180 );
				
				var newLayer = text_art_layer.duplicate();
				newLayer.rasterize(RasterizeType.ENTIRELAYER);
				var width = newLayer.bounds[2] - newLayer.bounds[0];
				newLayer.remove();
				
				current_height = this.convertToMillimeter( tuckbox.boxTexts()[i].height() ) - ( total_height / 2 ) + offset;
				
				var textPosition = [ x1 + ( width.as( "mm" ) / 2 ), 
									y_bottom - ( box_width_points * 0.8 ) + offset ];
				txtRef.position = textPosition;
				this.addStyleGlow( r, g, b, blendingMode, opacity, spread, size );
				
				text_art_layer = null;
				
				// LEFT
				text_art_layer = this._documentRef.artLayers.add();
				text_art_layer.name = "Text " + i + " - Left Layer";
				text_art_layer.kind = LayerKind.TEXT;
				
				txtRef = text_art_layer.textItem;
				txtRef.font = font;
				txtRef.contents = tuckbox.boxTexts()[i].text();
				txtRef.size = tuckbox.boxTexts()[i].height();
				txtRef.color = textColor;
				txtRef.kind = TextType.PARAGRAPHTEXT;
				
				text_art_layer.rotate( 270 );
				
				textPosition = [ x_left + ( box_depth_points * 0.5 ) - current_height, 
								y_bottom - ( box_width_points * 0.5 ) + ( width.as( "mm" ) / 2 ) ];
				txtRef.position = textPosition;
				this.addStyleGlow( r, g, b, blendingMode, opacity, spread, size );
				
				text_art_layer = null;
				
				// RIGHT
				text_art_layer = this._documentRef.artLayers.add();
				text_art_layer.name = "Text " + i + " - Right Layer";
				text_art_layer.kind = LayerKind.TEXT;
				
				txtRef = text_art_layer.textItem;
				txtRef.font = font;
				txtRef.contents = tuckbox.boxTexts()[i].text();
				txtRef.size = tuckbox.boxTexts()[i].height();
				txtRef.color = textColor;
				txtRef.kind = TextType.PARAGRAPHTEXT;
				
				text_art_layer.rotate( 90 );
				
				textPosition = [ x_right - ( box_depth_points * 0.5 ) + current_height, 
								y_bottom - ( box_width_points * 0.5 ) - ( width.as( "mm" ) / 2 ) ];
				txtRef.position = textPosition;
				this.addStyleGlow( r, g, b, blendingMode, opacity, spread, size );
				
				text_art_layer = null;
				
				// BOTTOM
				text_art_layer = this._documentRef.artLayers.add();
				text_art_layer.name = "Text " + i + " - Bottom Layer";
				text_art_layer.kind = LayerKind.TEXT;
				
				txtRef = text_art_layer.textItem;
				txtRef.font = font;
				txtRef.contents = tuckbox.boxTexts()[i].text();
				txtRef.size = tuckbox.boxTexts()[i].height();
				txtRef.color = textColor;
				txtRef.kind = TextType.PARAGRAPHTEXT;
				
				textPosition = [ x1 + ( box_width_points * 0.3 ) - ( width.as( "mm" ) / 2 ), 
								y_bottom - box_width_points - ( box_depth_points / 2 ) - current_height];
				txtRef.position = textPosition;
				this.addStyleGlow( r, g, b, blendingMode, opacity, spread, size );
				
				text_art_layer = null;
				
				// TOP
				text_art_layer = this._documentRef.artLayers.add();
				text_art_layer.name = "Text " + i + " - Top Layer";
				text_art_layer.kind = LayerKind.TEXT;
				
				txtRef = text_art_layer.textItem;
				txtRef.font = font;
				txtRef.contents = tuckbox.boxTexts()[i].text();
				txtRef.size = tuckbox.boxTexts()[i].height();
				txtRef.color = textColor;
				txtRef.kind = TextType.PARAGRAPHTEXT;
				
				text_art_layer.rotate( 180 );
				
				textPosition = [ x1 - ( box_width_points * 0.3 ) + ( width.as( "mm" ) / 2 ), 
								y_bottom - box_width_points - box_depth_points - box_width_points - ( box_depth_points / 2 ) + current_height ];
				txtRef.position = textPosition;
				this.addStyleGlow( r, g, b, blendingMode, opacity, spread, size );
				
				text_art_layer = null;
				
				offset += this.convertToMillimeter( tuckbox.boxTexts()[i].height() ) + linespace;
			}
			
			textColor = null;
		}
	},
	convertToMillimeter: function( points ) {
		var uv = new UnitValue( points + " pt" )
		return uv.as("mm");

	},
	convertToPoints: function( millimeters ) {
		var uv = new UnitValue( millimeters + " mm" )
		return uv.as("pt");
	},
	addStyleGlow: function( R, G, B, blendingMode, opacity, spread, size ) {
		var desc1 = new ActionDescriptor();
		var ref1 = new ActionReference();
		ref1.putProperty(this.cTID('Prpr'), this.cTID('Lefx'));
		ref1.putEnumerated(this.cTID('Lyr '), this.cTID('Ordn'), this.cTID('Trgt'));
		desc1.putReference(this.cTID('null'), ref1);
		var desc2 = new ActionDescriptor();
		desc2.putUnitDouble(this.cTID('Scl '), this.cTID('#Prc'), 100);
		// Glow color
		var desc4 = new ActionDescriptor();
		var rgb = new Array();
		desc4.putDouble(this.cTID('Rd  '), R);
		desc4.putDouble(this.cTID('Grn '), G);
		desc4.putDouble(this.cTID('Bl  '), B);
		// Blending mode of the effect
		var desc3 = new ActionDescriptor();
		desc3.putBoolean(this.cTID('enab'), true);
		desc3.putEnumerated( this.cTID('Md  '), this.cTID('BlnM'), this.cTID(blendingMode) );
		desc3.putObject(this.cTID('Clr '), this.sTID("RGBColor"), desc4);
		// Opacity
		desc3.putUnitDouble(this.cTID('Opct'), this.cTID('#Prc'), opacity);
		desc3.putEnumerated(this.cTID('GlwT'), this.cTID('BETE'), this.cTID('SfBL'));
		// Spread
		desc3.putUnitDouble(this.cTID('Ckmt'), this.cTID('#Pxl'), spread);
		// Size
		desc3.putUnitDouble(this.cTID('blur'), this.cTID('#Pxl'), size);
		// Noise
		desc3.putUnitDouble(this.cTID('Nose'), this.cTID('#Prc'), 0);
		// Quality: Jitter
		desc3.putUnitDouble(this.cTID('ShdN'), this.cTID('#Prc'), 0);
		desc3.putBoolean(this.cTID('AntA'), true);
		var desc5 = new ActionDescriptor();
		desc5.putString(this.cTID('Nm  '), "Linear");
		desc3.putObject(this.cTID('TrnS'), this.cTID('ShpC'), desc5);
		// Quality: Range
		desc3.putUnitDouble(this.cTID('Inpr'), this.cTID('#Prc'), 50);
		desc2.putObject(this.cTID('OrGl'), this.cTID('OrGl'), desc3);
		desc1.putObject(this.cTID('T   '), this.cTID('Lefx'), desc2);
		executeAction(this.cTID('setd'), desc1, DialogModes.NO);
	},
	cTID: function( s ) { 
		return app.charIDToTypeID( s ); 
	},
	sTID: function( s ) { 
		return app.stringIDToTypeID( s ); 
	},
	drawSinglePageBox: function( tuckbox ) {
		var box_outline_layer = this._documentRef.artLayers.add();
		box_outline_layer.name = "Box outline";

		var lineColor = new SolidColor();
		lineColor.rgb.red = 255;
		lineColor.rgb.green = 255;
		lineColor.rgb.blue = 255;
		app.foregroundColor = lineColor;
		
		var x_left = Math.ceil( ( this._documentRef.width.as( "pt" ) - tuckbox.totalWidth().as( "pt" ) ) / 2 );
		var x_right = x_left + tuckbox.totalWidth().as( "pt" );
		var y_top = Math.ceil( ( this._documentRef.height.as( "pt" ) - tuckbox.totalHeight().as( "pt" ) ) / 2 );
		var y_bottom = y_top + tuckbox.totalHeight().as( "pt" );
		
		var temp5 = UnitValue( 5, "mm" );
		var temp10 = UnitValue( 10, "mm" );
		var temp7_5 = UnitValue( 7.5, "mm" );
		
		// Bottom line
		var x1 = x_left;
		var x2 = ( x_left + x_right ) / 2 - temp5.as( "pt" );
		var y1 = y_bottom;
		var y2 = y1;
		this.drawLine( [x1, y1], [x2, y2] );
		
		x1 = x2 + temp10.as( "pt" );
		this.drawCurvedLine( [x1, y_bottom], [x2, y_bottom], [x1, y_bottom - temp7_5.as( "pt" )], [x2, y_bottom - temp7_5.as( "pt" )] );
		
		x2 = x_right;
		this.drawLine( [x1, y1], [x2, y2] );
		
		// second horizontal line from the bottom
		y1 = y1 -= tuckbox.width().as( "pt" );
		y2 = y1;
		x1 = x_left;
		x2 = x1 + tuckbox.depth().as( "pt" );
		this.drawLine( [x1, y1], [x2, y2] );
		
		x1 = x2;
		x2 = x1 + tuckbox.height().as( "pt" );
		this.drawDashedLine( [x1, y1], [x2, y2] );
		
		x1 = x2;
		x2 = x_right;
		this.drawLine( [x1, y1], [x2, y2] );
		
		// third horizontal line from the bottom
		y1 = y1 -= tuckbox.depth().as( "pt" );
		y2 = y1;	
		x1 = x_left;
		this.drawDashedLine( [x1, y1], [x2, y2] );
		
		// fourth horizontal line from the bottom
		y1 = y1 -= tuckbox.width().as( "pt" );
		y2 = y1;	
		this.drawDashedLine( [x1, y1], [x2, y2] );
		
		if( tuckbox.depth().as( "pt" ) <= ( tuckbox.height().as( "pt" ) / 2 ) )
		{
			// left vertical line from the bottom
			x2 = x_left;
			y2 = y_bottom;
			this.drawLine( [x1, y1], [x2, y2] );
			
			// right vertical line from the bottom
			x1 = x_right;
			x2 = x_right;
			this.drawLine( [x1, y1], [x2, y2] );
		}
		else
		{
			var y_temp = y1;
			
			// left vertical line from the top
			x2 = x_left;
			y2 = y_bottom - tuckbox.width().as( "pt" ) - tuckbox.depth().as( "pt" ) + ( tuckbox.height().as( "pt" ) * 0.5 );
			this.drawLine( [x1, y1], [x2, y2] );
			
			x1 = x2 + tuckbox.depth().as( "pt" );
			this.drawLine( [x1, y2], [x2, y2] );
			
			// right vertical line from the top
			x1 = x_right;
			x2 = x_right;
			this.drawLine( [x1, y1], [x2, y2] );
			
			x2 = x1 - tuckbox.depth().as( "pt" );
			this.drawLine( [x1, y2], [x2, y2] );
			
			x1 = x_right;
			x2 = x_right;
			y1 = y_bottom;
			y2 = y_bottom - tuckbox.width().as( "pt" );
			this.drawLine( [x1, y1], [x2, y2] );
			
			x1 = x_left;
			x2 = x_left;
			y2 = y_bottom - tuckbox.width().as( "pt" );
			this.drawLine( [x1, y1], [x2, y2] );
			
			y1 = y_temp;
		}
		
		x1 = x_left + tuckbox.depth().as( "pt" );
		x2 = x1 + tuckbox.height().as( "pt" );
		y2 = y1 - tuckbox.depth().as( "pt" );
		this.drawLine( [x1, y1], [x1, y2] );
		this.drawDashedLine( [x1, y2], [x2, y2] );
		this.drawLine( [x2, y1], [x2, y2] );
		
		y2 = y1 + tuckbox.width().as( "pt" );
		this.drawDashedLine( [x1, y1], [x1, y2] );
		this.drawDashedLine( [x2, y1], [x2, y2] );
		
		y1 = y2 + tuckbox.depth().as( "pt" );
		this.drawLine( [x1, y1], [x1, y2] );
		this.drawLine( [x2, y1], [x2, y2] );
		
		y2 = y_bottom;
		this.drawDashedLine( [x1, y1], [x1, y2] );
		this.drawDashedLine( [x2, y1], [x2, y2] );
		
		x1 = x_left;
		x2 = x1 + tuckbox.depth().as( "pt" );
		y1 = y_bottom - ( tuckbox.width().as( "pt" ) + tuckbox.depth().as( "pt" ) + tuckbox.width().as( "pt" ) );
		y2 = y1 - tuckbox.depth().as( "pt" );
		this.drawCurvedLine( [x1, y1], [x2, y2], [x1, (y1 + y2) / 2], [(x1 + x2) / 2, y2] );
		
		x1 = x2;
		x2 = x1 + tuckbox.lip().as( "pt" );
		y1 = y2;
		y2 = y1 - tuckbox.lip().as( "pt" );
		this.drawCurvedLine( [x1, y1], [x2, y2], [x1, (y1 + y2) / 2], [(x1 + x2) / 2, y2] );
		
		x1 = x2 + tuckbox.height().as( "pt" ) - tuckbox.lip().as( "pt" ) - tuckbox.lip().as( "pt" );
		this.drawLine( [x1, y2], [x2, y2] );
		
		x2 = x1;
		x1 = x_right - tuckbox.depth().as( "pt" );
		this.drawCurvedLine( [x1, y1], [x2, y2], [x1, (y1 + y2) / 2], [(x1 + x2) / 2, y2] );
		
		x1 = x_right;
		x2 = x1 - tuckbox.depth().as( "pt" );
		y1 = y_bottom - ( tuckbox.width().as( "pt" ) + tuckbox.depth().as( "pt" ) + tuckbox.width().as( "pt" ) );
		y2 = y1 - tuckbox.depth().as( "pt" );
		this.drawCurvedLine( [x1, y1], [x2, y2], [x1, (y1 + y2) / 2], [(x1 + x2) / 2, y2] );
		
		lineColor = null;
		
		box_outline_layer = null;
	},
	drawLine: function( start, stop ) {	
		var startPoint = new PathPointInfo();
		startPoint.anchor = start;
		startPoint.leftDirection = start;
		startPoint.rightDirection = start;
		startPoint.kind = PointKind.CORNERPOINT;
		
		var stopPoint = new PathPointInfo();
		stopPoint.anchor = stop;
		stopPoint.leftDirection = stop;
		stopPoint.rightDirection = stop;
		stopPoint.kind = PointKind.CORNERPOINT;
		
		var spi = new SubPathInfo();
		spi.closed = false;
		spi.operation = ShapeOperation.SHAPEXOR;
		spi.entireSubPath = [startPoint, stopPoint];
		
		var line = this._documentRef.pathItems.add("Line", [spi]);
		line.strokePath(ToolType.PENCIL);
		line.remove();
	},drawCurvedLine: function( start, stop, left, right ) {
		var startPoint = new PathPointInfo();
		startPoint.anchor = start;
		startPoint.leftDirection = left;
		startPoint.rightDirection = start;
		startPoint.kind = PointKind.CORNERPOINT;
		
		var stopPoint = new PathPointInfo();
		stopPoint.anchor = stop;
		stopPoint.leftDirection = stop;
		stopPoint.rightDirection = right;
		stopPoint.kind = PointKind.CORNERPOINT;
		
		var spi = new SubPathInfo();
		spi.closed = false;
		spi.operation = ShapeOperation.SHAPEXOR;
		spi.entireSubPath = [startPoint, stopPoint];
		
		var line = this._documentRef.pathItems.add("Line", [spi]);
		line.strokePath(ToolType.PENCIL);
		line.remove();
	},
	drawDashedLine: function( start, stop ) {
		var lineColor = new SolidColor();
		lineColor.rgb.red = 128;
		lineColor.rgb.green = 128;
		lineColor.rgb.blue = 128;
		app.foregroundColor = lineColor;
		
		var dash_length = this.convertToPoints( 1 );
		var start_vector = new Vector( start[0], start[1] );
		var stop_vector = new Vector( stop[0], stop[1] );
		var direction_vector = new Vector( stop[0] - start[0], stop[1] - start[1] );
		var length = direction_vector.length();
		var direction_unit_vector = new Vector( direction_vector.x / length, direction_vector.y / length );
		
		var sections = Math.ceil( length / ( dash_length * 2 ) );
		var section_start = 0;
		var section_stop = 0;
		
		var i = 0;
		for( i = 0; i < sections; i++ )
		{
			section_start_x = start_vector.x + (direction_unit_vector.x * dash_length * 2 * i);
			section_start_y = start_vector.y + (direction_unit_vector.y * dash_length * 2 * i);
			
			section_stop_x = Math.min( stop[0], section_start_x + (direction_unit_vector.x * dash_length) );
			section_stop_y = Math.min( stop[1], section_start_y + (direction_unit_vector.y * dash_length) );
			
			section_start = [section_start_x, section_start_y];
			section_stop = [section_stop_x, section_stop_y];
			this.drawLine( section_start, section_stop );
		}
		
		var lineColor = new SolidColor();
		lineColor.rgb.red = 255;
		lineColor.rgb.green = 255;
		lineColor.rgb.blue = 255;
		app.foregroundColor = lineColor;
		
		lineColor = null;
	},
	drawPageInfo: function( infoTexts, cardAmount ) {
		var font = "Arial";
		
		if( infoTexts.length > 0 )
		{		
			var textColor = new SolidColor();
			textColor.rgb.red = 0;
			textColor.rgb.green = 0;
			textColor.rgb.blue = 0;
			
			var linespace = 2;
			
			var i = 0;
			for( i = infoTexts.length - 1; i >= 0; i-- )
			{			
				var text_art_layer = this._documentRef.artLayers.add();
				text_art_layer.name = "Info text Layer " + i;
				text_art_layer.kind = LayerKind.TEXT;
				
				var txtRef = text_art_layer.textItem;
				txtRef.font = font;
				txtRef.contents = infoTexts[i];
				txtRef.size = 12;
				txtRef.color = textColor;
				txtRef.kind = TextType.PARAGRAPHTEXT;
				
				textPosition = [ 10, 10 + ( i * this.convertToMillimeter( 12 + linespace ) ) ];
				txtRef.position = textPosition;
				
				text_art_layer = null;
			}
		}
		
		var texts = infoTexts.length;
		
		var text_art_layer = this._documentRef.artLayers.add();
		text_art_layer.name = "Info text Layer " + texts;
		text_art_layer.kind = LayerKind.TEXT;
		
		var txtRef = text_art_layer.textItem;
		txtRef.font = font;
		txtRef.contents = "Holds " + cardAmount + " cards";
		txtRef.size = 12;
		txtRef.color = textColor;
		txtRef.kind = TextType.PARAGRAPHTEXT;
		
		textPosition = [ 10, 10 + ( texts * this.convertToMillimeter( 12 + linespace ) ) ];
		txtRef.position = textPosition;
		
		text_art_layer = null;
		
		texts++;
		texts++;
		var text_art_layer = this._documentRef.artLayers.add();
		text_art_layer.name = "Info text Layer " + ( texts - 1 );
		text_art_layer.kind = LayerKind.TEXT;
		
		var txtRef = text_art_layer.textItem;
		txtRef.font = font;
		txtRef.contents = "By mephisto0666";
		txtRef.size = 12;
		txtRef.color = textColor;
		txtRef.kind = TextType.PARAGRAPHTEXT;
		
		textPosition = [ 10, 10 + ( texts * this.convertToMillimeter( 12 + linespace ) ) ];
		txtRef.position = textPosition;
		
		text_art_layer = null;
	},
	cleanup: function() {
		this._imageDocRef = null;
		this._fileRef = null;
		this._documentRef = null;
	},
	restoreSettings: function() {
		// Restore original ruler unit setting
		preferences.rulerUnits = this._defaultRulerUnits;
		UnitValue.baseUnit = null; //restore default
	}
};