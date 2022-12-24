#include 'TuckboxGenerator.jsx'

var kingdomDeathMonsterTuckBoxImage = "D:\\Temp\\Tuckboxes\\KDM\\kdm.png";
var kingdomDeathMonsterLogoImage = "D:\\Temp\\Tuckboxes\\KDM\\logo.png";
// Kingdom Death: Monster: White Lion Cards
var kingdomDeathMonsterWhiteLionTuckBox = new TuckBox( 	
	76, 
	[new BoxText("White Lion Cards", 16)], 
	kingdomDeathMonsterTuckBoxImage, 
	["D:\\Temp\\Tuckboxes\\KDM\\whitelion_basic.png",
	"D:\\Temp\\Tuckboxes\\KDM\\whitelion_ai.png",
	"D:\\Temp\\Tuckboxes\\KDM\\whitelion_hitlocation.png",
	"D:\\Temp\\Tuckboxes\\KDM\\whitelion_hunt.png",
	"D:\\Temp\\Tuckboxes\\KDM\\whitelion_resource.png"], 
	kingdomDeathMonsterLogoImage, 
	["Horizontal Tuckboxes: Kingdom Death: Monster", 
	"White Lion v 1.0"], 
	ffgRegularSleevedCard, 
	"Kingdom Death: Monster White Lion Cards" );
	
var tuckboxDrawer = new TuckboxDrawer( dpi, documentSize.LETTER );
tuckboxDrawer.drawTuckbox( kingdomDeathMonsterWhiteLionTuckBox );
tuckboxDrawer = null;