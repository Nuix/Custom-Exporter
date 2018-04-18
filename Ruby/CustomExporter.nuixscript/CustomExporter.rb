# Menu Title: Custom Exporter
# Needs Case: true
# Needs Selected Items: false

# Bootstrap the Nx library
script_directory = File.dirname(__FILE__)
require File.join(script_directory,"Nx.jar")
java_import "com.nuix.nx.NuixConnection"
java_import "com.nuix.nx.LookAndFeelHelper"
java_import "com.nuix.nx.dialogs.ChoiceDialog"
java_import "com.nuix.nx.dialogs.TabbedCustomDialog"
java_import "com.nuix.nx.dialogs.CommonDialogs"
java_import "com.nuix.nx.dialogs.ProgressDialog"
java_import "com.nuix.nx.misc.PlaceholderResolver"
java_import "com.nuix.nx.helpers.MetadataProfileHelper"
java_import "com.nuix.nx.controls.models.Choice"

LookAndFeelHelper.setWindowsIfMetal
NuixConnection.setUtilities($utilities)
NuixConnection.setCurrentNuixVersion(NUIX_VERSION)

# Load class for working with DAT file
load File.join(script_directory,"DAT.rb_")

# Load class for exporting XLSX
load File.join(script_directory,"Xlsx.rb")

# Require CSV library
require 'csv'

# Get a listing of metadata profiles that contain at least the field GUID a
# (requirement for the script to make use of them), make sure we end up
# with at least one for the user to select from
metadata_profiles = $utilities.getMetadataProfileStore.getMetadataProfiles
metadata_profiles = metadata_profiles.select{|p|MetadataProfileHelper.profileContainsField("GUID",p)}
profile_names = metadata_profiles.map{|p|p.getName}
if profile_names.size < 1
	CommonDialogs.showError("This script requires that at least one metadata profile that contains the field 'GUID' exist.")
	exit 1
end

# Get list of custom metadata field names
custom_field_names = $current_case.getCustomMetadataFields("user")

# Get list of production set names
production_set_names = $current_case.getProductionSets.map{|ps|ps.getName}

#=======================#
# Build Settings Dialog #
#=======================#
dialog = TabbedCustomDialog.new("Custom Exporter")
dialog.enableStickySettings(File.join(script_directory,"RecentSettings.json"))
dialog.setHelpFile(File.join(script_directory,"Help.html"))

# Main Tab
main_tab = dialog.addTab("main_tab","Main")
main_tab.appendDirectoryChooser("export_directory","Export Directory")
main_tab.appendCheckBox("export_csv","Export Metadata As CSV",false)
main_tab.appendCheckBox("export_xlsx","Export Metadata As XLSX",false)
main_tab.appendLabel("loadfile_note","Note: You will always get a DAT file exported.")
main_tab.appendHeader("Note: Only profiles containing 'GUID' are listed.")
main_tab.appendComboBox("metadata_profile","Metadata Profile",profile_names)

main_tab.appendSeparator("Input Items")
main_tab.appendCheckBox("include_families","Include Families",false)
main_tab.appendRadioButton("use_query","Use Query","item_source",true)
if $current_selected_items.size > 0
	main_tab.appendRadioButton("use_selected_items","Use Selected Items: #{$current_selected_items.size}","item_source",false)
end
if production_set_names.size > 0
	main_tab.appendRadioButton("use_production_set","Use Production Set","item_source",false)
	main_tab.appendComboBox("source_production_set","Production Set",production_set_names)
	main_tab.enabledOnlyWhenChecked("source_production_set","use_production_set")
end
main_tab.appendHeader("Item Query")
main_tab.appendTextArea("source_query","","")
main_tab.enabledOnlyWhenChecked("source_query","use_query")
main_tab.enabledOnlyWhenChecked("include_families","use_query")

# Text Settings Tab
text_tab = dialog.addTab("text_tab","Text")
text_tab.appendCheckBox("export_text","Export Text",true)
text_tab.appendTextField("text_template","Text Path Template","{export_directory}\\TEXT\\{box_major}\\{box}\\{name}.txt")
text_tab.enabledOnlyWhenChecked("text_template","export_text")

# PDF settings Tab
pdf_tab = dialog.addTab("pdf_tab","PDF")
pdf_tab.appendCheckBox("export_pdf","Export PDFs",true)
pdf_tab.appendTextField("pdf_template","PDF Path Template","{export_directory}\\PDF\\{box_major}\\{box}\\{name}.pdf")
pdf_tab.enabledOnlyWhenChecked("pdf_template","export_pdf")

# TIFF settings Tab
tiff_tab = dialog.addTab("tiff_tab","TIFF")
tiff_tab.appendCheckBox("export_tiff","Export TIFFs",true)
tiff_tab.appendTextField("tiff_template","TIFF Path Template","{export_directory}\\TIFF\\{box_major}\\{box}\\{name}.tiff")
dpi_choices = [
	#"75", # Is this still supported actually?  Gave me an error
	"150",
	"300"
]
tiff_tab.appendComboBox("tiff_dpi","TIFF DPI",dpi_choices)
tiff_formats = [
	"MONOCHROME CCITT T6 G4",
	"GREYSCALE_UNCOMPRESSED",
	"GREYSCALE_DEFLATE",
	"GREYSCALE_LZW",
	"COLOUR UNCOMPRESSED",
	"COLOUR DEFLATE",
	"COLOUR LZW"
]
tiff_tab.appendComboBox("tiff_format","TIFF Format",tiff_formats)
tiff_tab.enabledOnlyWhenChecked("tiff_template","export_tiff")
tiff_tab.enabledOnlyWhenChecked("tiff_dpi","export_tiff")

# Native settings Tab
native_tab = dialog.addTab("native_tab","Natives")
native_tab.appendCheckBox("export_natives","Export Natives",true)
native_tab.appendTextField("natives_template","Natives Path Template","{export_directory}\\NATIVE\\{box}\\{name}.{extension}")
native_tab.appendComboBox("natives_email_format","Email Format",["msg","eml","html"])
native_tab.appendCheckBox("include_attachments","Include Attachments on Emails",true)
native_tab.enabledOnlyWhenChecked("natives_template","export_natives")
native_tab.enabledOnlyWhenChecked("natives_email_format","export_natives")

# Placeholder settings tab
placeholders_tab = dialog.addTab("placeholders_tab","Placeholders")

placeholders_tab.appendSeparator("Box Minor Settings")
placeholders_tab.appendSpinner("box_start","{box} starting number",0,0,99999999)
placeholders_tab.appendSpinner("box_width","{box} zero fill width",4,1,8)
placeholders_tab.appendSpinner("box_step","Items per {box} increment",1000,1,10000000,10)

placeholders_tab.appendSeparator("Box Major Settings")
placeholders_tab.appendSpinner("box_major_width","{box_major} zero fill width",4,1,8)

placeholders_tab.appendSeparator("Production Set Settings")
if production_set_names.size > 0
	placeholders_tab.appendCheckBox("enable_docid","Enable {docid} placeholder",false)
	placeholders_tab.appendComboBox("docid_prod_set","Source Production Set",production_set_names)
	placeholders_tab.enabledOnlyWhenChecked("docid_prod_set","enable_docid")
else
	placeholders_tab.appendHeader("Current case has no production sets defined.")
end

# These are to allow the user to provide fixed placeholders.  Use case is when there is a standard
# structure but pehaps each export has one value change.  User can include a user value placeholder
# wherever its needed but they will only need to change the value here on subsequent runs.
placeholders_tab.appendSeparator("User Placeholders")
placeholders_tab.appendTextField("user_value_1","{user_1} = ","")
placeholders_tab.appendTextField("user_value_2","{user_2} = ","")
placeholders_tab.appendTextField("user_value_3","{user_3} = ","")
placeholders_tab.appendTextField("user_value_4","{user_4} = ","")
placeholders_tab.appendTextField("user_value_5","{user_5} = ","")

# Only have option for custom field names if the case actually has some
# custom metadata fields present
placeholders_tab.appendSeparator("Custom Metadata")
if custom_field_names.size > 0
	placeholders_tab.appendCheckBox("use_custom_placeholders","Use custom metadata placeholders",false)
	placeholders_tab.appendComboBox("custom_field_1","{custom_1} = ",custom_field_names)
	placeholders_tab.appendComboBox("custom_field_2","{custom_2} = ",custom_field_names)
	placeholders_tab.appendComboBox("custom_field_3","{custom_3} = ",custom_field_names)
	placeholders_tab.appendComboBox("custom_field_4","{custom_4} = ",custom_field_names)
	placeholders_tab.appendComboBox("custom_field_5","{custom_5} = ",custom_field_names)
	placeholders_tab.enabledOnlyWhenChecked("custom_field_1","use_custom_placeholders")
	placeholders_tab.enabledOnlyWhenChecked("custom_field_2","use_custom_placeholders")
	placeholders_tab.enabledOnlyWhenChecked("custom_field_3","use_custom_placeholders")
	placeholders_tab.enabledOnlyWhenChecked("custom_field_4","use_custom_placeholders")
	placeholders_tab.enabledOnlyWhenChecked("custom_field_5","use_custom_placeholders")
else
	placeholders_tab.appendHeader("Current case has no custom metadata fields defined.")
end

# Allows use to select tags placeholder
tags_placeholder_tab = dialog.addTab("tags_placeholder_tab","{tags} Placeholder")
all_tag_choices = $current_case.getAllTags.map{|tag| Choice.new(tag,tag,tag,true)}
tags_placeholder_tab.appendChoiceTable("placeholder_tags","Allowed Placeholder Tags",all_tag_choices)

# Path filtering settings Tab
filtered_path_tab = dialog.addTab("filtered_path_tab","Filtered Path Name")
filtered_path_tab.appendCheckBox("filter_dat_item_path","Modify 'Path Name' in DAT if Present",false)
type_choices = $utilities.getItemTypeUtility.getAllTypes.sort_by{|t| [t.getKind.getName,t.getLocalisedName,t.getName]}
type_choices = type_choices.map{|t|Choice.new(t.getName,"#{t.getKind} | #{t.getLocalisedName} | #{t.getName}")}
filtered_path_tab.appendHeader("Check types which will not be included in {filtered_item_path} and optionally removed from 'Path Name' in the DAT.")
filtered_path_tab.appendChoiceTable("filtered_path_mime_types","Mime Types",type_choices)

# Worker Settings Tab
worker_tab = dialog.addTab("worker_tab","Worker Settings")
worker_tab.appendLocalWorkerSettings("worker_settings")

# Provide a callback to the settings dialog which will validate settings when the user
# clicks the ok button.  If this callback yields false, settings dialog will not close
# if yields true dialog will close and dialog result will be true when retrieved later
dialog.validateBeforeClosing do |values|
	# Validate export directory
	if values["export_directory"].nil? || values["export_directory"].strip.empty?
		CommonDialogs.showError("Export Directory cannot be empty.")
		next false
	end

	# Validate all template paths are not empty
	if values["export_text"] && (values["text_template"].nil? || values["text_template"].strip.empty?)
		CommonDialogs.showError("Text Path Template cannot be empty.")
		next false
	end
	if values["export_natives"] && (values["natives_template"].nil? || values["natives_template"].strip.empty?)
		CommonDialogs.showError("Natives Path Template cannot be empty.")
		next false
	end
	if values["export_pdf"] && (values["pdf_template"].nil? || values["pdf_template"].strip.empty?)
		CommonDialogs.showError("PDF Path Template cannot be empty.")
		next false
	end
	if values["export_tiff"] && (values["tiff_template"].nil? || values["tiff_template"].strip.empty?)
		CommonDialogs.showError("TIFF Path Template cannot be empty.")
		next false
	end

	# Validate worker temp
	if values["worker_settings"]["workerTemp"].strip.empty?
		CommonDialogs.showError("Please provide a path for worker temp in on 'Worker Settings' tab.")
		next false
	end

	# Validate metadata profile contains GUID
	selected_profile = $utilities.getMetadataProfileStore.getMetadataProfile(values["metadata_profile"])
	if !selected_profile.getMetadata.any?{|field|field.getName == "GUID"}
		message = "Selected metadata profile is required to include 'GUID'."
		message << "\nThis is needed to correlate entries in DAT back to actual items in the case."
		CommonDialogs.showError(message)
		next false
	end

	# If no validations failed we yield true to signal we're good to proceed
	next true
end

# Method used later on to move an export product file to appropriate structure
# based on template user provided
def restructure_product(record,product_path_field,product_name,resolver,current_item,template,temp_export_directory,export_directory,placeholder_tags)
	original_path = record[product_path_field]
	original_fullpath = File.join(temp_export_directory,original_path).gsub(/\//,"\\")

	pathing_data = []

	if !java.io.File.new(original_fullpath).exists
		pd.logMessage("Could not find #{product_name} file: #{original_fullpath}")
	else
		current_item_tags = current_item.getTags.select{|tag| placeholder_tags[tag] == true}
		current_item_tags << "No Tags" if current_item_tags.size < 1
		current_item_tags.each do |tag|
			tag_path = tag.gsub(/\|/,"\\")
			resolver.set("tags",tag_path)
			updated_fullpath = resolver.resolveTemplatePath(template)
			updated_path = updated_fullpath.gsub(/^#{Regexp.escape(export_directory)}\\/i,".\\")
			pathing_data << {
				:updated_fullpath => updated_fullpath,
				:updated_path => updated_path
			}
		end

		pathing_data.uniq!
		
		had_error = false
		pathing_data.each_with_index do |path_entry,path_entry_index|
			updated_fullpath = path_entry[:updated_fullpath]
			updated_path = path_entry[:updated_path]
			java.io.File.new(updated_fullpath).getParentFile.mkdirs
			begin
				puts "Renaming #{original_fullpath} to #{updated_fullpath}"
				if path_entry_index == pathing_data.size - 1
					java.io.File.new(original_fullpath).renameTo(java.io.File.new(updated_fullpath))
				else
					org.apache.commons.io.FileUtils.copyFile(java.io.File.new(original_fullpath),java.io.File.new(updated_fullpath))
				end
			rescue Exception => exc
				pd.logMessage("Error while renaming #{product_name} for #{current_item.getGuid}:")
				pd.logMessage(exc.message)
				record[product_path_field] = exc.message
				had_error = true
			end
		end
		if !had_error
			record[product_path_field] = pathing_data.map{|path_entry| path_entry[:updated_path]}.join("; ")
		end
	end
end

# Display dialog
dialog.display
if dialog.getDialogResult == true
	# Do the work and show a progress dialog to provide the user feedback
	ProgressDialog.forBlock do |pd|
		pd.setAbortButtonVisible(false)
		pd.setSubProgressVisible(false)
		pd.setTitle("Custom Exporter")

		# Get settings from settings dialog
		values = dialog.toMap
		# Convenience handles to item sorter and item utility
		sorter = $utilities.getItemSorter
		iutil = $utilities.getItemUtility

		# Obtain the items we will be working with depending on the settings selected
		pd.setMainStatusAndLogIt("Getting input items...")
		items = nil
		source_production_set = nil
		if values["use_selected_items"]
			items = $current_selected_items.to_a
		elsif values["use_query"]
			pd.logMessage("Searching: #{values["source_query"]}")
			items = $current_case.search(values["source_query"])
		elsif values["use_production_set"]
			source_production_set = $current_case.getProductionSets.select{|ps|ps.getName == values["source_production_set"]}.first
			pd.logMessage("Getting items from production set: #{values["source_production_set"]}")
			items = source_production_set.getItems
		end

		if values["include_families"]
			pd.logMessage("Including family items...")
			items = iutil.findFamilies(items)
		end

		pd.logMessage("Removing any excluded items...")
		items = items.reject{|i|i.isExcluded}

		# Remove evidence items
		pd.logMessage("Removing any evidence container items...")
		items = items.reject{|i|i.getType.getName == "application/vnd.nuix-evidence"}

		pd.logMessage("Input Items: #{items.size}")

		# Extract settings dialog values into variables for convenience later
		export_text = values["export_text"]
		text_template = values["text_template"]
		export_natives = values["export_natives"]
		natives_template = values["natives_template"]
		natives_email_format = values["natives_email_format"]
		export_pdf = values["export_pdf"]
		pdf_template = values["pdf_template"]
		export_tiff = values["export_tiff"]
		tiff_template = values["tiff_template"]

		export_directory = values["export_directory"]
		temp_export_directory = File.join(export_directory,"TEMP")
		use_custom_placeholders = values["use_custom_placeholders"]
		custom_field_1 = values["custom_field_1"]
		custom_field_2 = values["custom_field_2"]
		custom_field_3 = values["custom_field_3"]
		custom_field_4 = values["custom_field_4"]
		custom_field_5 = values["custom_field_5"]
		user_value_1 = values["user_value_1"]
		user_value_2 = values["user_value_2"]
		user_value_3 = values["user_value_3"]
		user_value_4 = values["user_value_4"]
		user_value_5 = values["user_value_5"]
		enable_docid = values["enable_docid"]

		export_csv = values["export_csv"]
		export_xlsx = values["export_xlsx"]
		
		docid_prod_set = nil
		if enable_docid
			docid_prod_set = $current_case.getProductionSets.select{|ps|ps.getName == values["docid_prod_set"]}.first
		end
		filter_dat_item_path = values["filter_dat_item_path"]

		# Calculate whether we will need to calculate filtered item path for each item as
		# this operation could be costly regarding performance
		fp_regex = /\{filtered_item_path\}/i
		generate_filtered_path = filter_dat_item_path || (export_text && text_template =~ fp_regex) ||
			(export_natives && natives_template =~ fp_regex) || (export_pdf && pdf_template =~ fp_regex) ||
			user_value_1 =~ fp_regex || user_value_2 =~ fp_regex || user_value_3 =~ fp_regex ||
			user_value_4 =~ fp_regex || user_value_5 =~ fp_regex
		filtered_path_mime_types = {}
		values["filtered_path_mime_types"].each{|mt|filtered_path_mime_types[mt] = true}

		# Build DOCID lookup if needed
		docid_lookup = Hash.new{|h,k|h[k]=""}
		if enable_docid
			pd.setMainStatusAndLogIt("Building DOCID cross reference...")
			prod_items = docid_prod_set.getProductionSetItems
			prod_items.each do |prod_item|
				docid_lookup[prod_item.getItem] = prod_item.getDocumentNumber.toString
			end
		end

		# Pull tags user selected in settings dialog into a hash for quick lookup later
		placeholder_tags = {}
		values["placeholder_tags"].each{|tag| placeholder_tags[tag] = true}

		if values["close_tabs"]
			$window.closeAllTabs
		end

		# Build the batch exporter.  We will be using this to perform an initial temporary export
		# which we will then as a secondary step restructure based on the settings provided
		pd.setMainStatusAndLogIt("Configuring exporter...")
		exporter = $utilities.createBatchExporter(temp_export_directory)

		pd.setMainStatusAndLogIt("\tConfiguring worker settings...")
		exporter.setParallelProcessingSettings(values["worker_settings"])

		# Note that script uses DAT file from initial export during the restructuring step
		# so exporting a concordance loadfile is required!
		pd.setMainStatusAndLogIt("\tAdding loadfile...")
		exporter.addLoadFile("concordance",{:metadataProfile => values["metadata_profile"]})

		# Configure to export text if settings specify this
		if export_text
			pd.setMainStatusAndLogIt("\tAdding text...")
			text_settings = {
				:naming => "guid",
				:path => "TEXT",
			}
			exporter.addProduct("text",text_settings)
		end

		# Configure to export natives if settings specify this
		if export_natives
			pd.setMainStatusAndLogIt("\tAdding natives...")
			natives_settings = {
				:naming => "guid",
				:path => "NATIVE",
				:mailFormat => natives_email_format,
				:includeAttachments => values["include_attachments"],
			}
			exporter.addProduct("native",natives_settings)
		end

		# Configure to export PDFs if settings specify this
		if export_pdf
			pd.setMainStatusAndLogIt("\tAdding PDFs...")
			pdf_settings = {
				:naming => "guid",
				:path => "PDF",
			}
			exporter.addProduct("pdf",pdf_settings)
		end

		# Configure to export TIFFs if settings specify this
		if export_tiff
			pd.setMainStatusAndLogIt("\tAdding TIFFs...")
			tiff_settings = {
				:naming => "guid",
				:path => "TIFF",
				:multiPageTiff => true,
				:tiffDpi => values["tiff_dpi"].to_i,
				:tiffFormat => values["tiff_format"].gsub(" ","_")
			}
			exporter.addProduct("tiff",tiff_settings)
		end

		# Get progress dialog ready and hookup export callback so that it will
		# update the progress dialog
		pd.setMainProgress(0,items.size)
		current_export_stage = nil
		exporter.whenItemEventOccurs do |info|
			stage_name = info.getStage
			pd.setSubStatus("Export Stage: #{stage_name}")
			pd.setMainProgress(info.getStageCount)
			if current_export_stage != stage_name
				pd.logMessage("\tExport Stage: #{stage_name}")
				current_export_stage = stage_name
			end
		end

		#Can only call this if the the licence has the feature "PRODUCTION_SET"
		if $utilities.getLicence.hasFeature("PRODUCTION_SET")
			exporter.setNumberingOptions({"createProductionSet" => false})
		end

		# We are ready to begin exporting!
		pd.setMainStatusAndLogIt("Performing pre-customization export...")
		if values["use_production_set"]
			exporter.exportItems(source_production_set)
		else
			exporter.exportItems(items)
		end
		pd.setSubStatus("")
		pd.logMessage("Export completed")
		pd.setMainProgress(0,1)

		# Exporting should be complete when we get here, so we are ready to take the initial
		# tempory export and restructure the file names and path based on templates provided
		# by the user, we are also going to be making sure we produce a DAT file with the 
		# product paths updated to match
		exported_temp_dat = File.join(temp_export_directory,"loadfile.dat")
		final_dat = File.join(export_directory,"loadfile.dat")
		resolver = PlaceholderResolver.new
		record_number = 0

		box_start = values["box_start"]
		box_width = values["box_width"]
		box_step = values["box_step"]

		box_major_width = values["box_major_width"]

		items_by_guid = {}
		items.each{|item|items_by_guid[item.getGuid]=item}
		pd.setMainProgress(0,items.size)

		# Used to periodically report progress updates
		last_progress = Time.now

		# Initialize other loadfile export stuff if needed
		csv_file = File.join(export_directory,"loadfile.csv")
		xlsx_file = File.join(export_directory,"loadfile.xlsx")

		xlsx = nil
		sheet = nil
		csv = nil

		if export_csv
			csv = CSV.open(csv_file,"w:utf-8")
		end

		if export_xlsx
			xlsx = Xlsx.new(xlsx_file)
			sheet = xlsx.get_sheet("Load File")
		end

		# Restructuring process is driven by this transpose process.  What happens is that DAT.transpose_each
		# will read the initial temporary exported DAT line by line, yielding to the block a hash for each record
		# where the key is the column name and the value is the value for that given record.  The path values are
		# calculated from the templates the user provided.  Using the originally exported path from the DAT and the
		# newly calculated path the exported product is renamed (file system move really).  Finally DAT.transpose_each
		# converts the now modified record hash to a DAT formatted record line and writes that to the final DAT file.
		pd.setMainStatusAndLogIt("Customizing exported data...")
		DAT.transpose_each(exported_temp_dat,final_dat) do |record|
			record_number += 1
			if (Time.now - last_progress) > 1
				pd.setMainProgress(record_number)
				pd.setSubStatus("#{record_number}/#{items.size}")
				last_progress = Time.now
			end
			
			# Here we are generating a box number.  Essentially we calculate the box number
			# as incrementing every N items.  When then convert the number to a string zero padded
			# to the total length of box width and box major width.  Box gets the right half, box major
			# gets the left half.  Some additional formatting manipulations are then performed to ensure
			# 0 fill is correct in each piece while allowing box major to overflow its fill width if necessary.
			box_number = ((record_number.to_f / box_step.to_f).floor + box_start).to_i
			box_whole_string = box_number.to_s.rjust(box_width+box_major_width,"0")
			box_len = box_whole_string.size
			box_min = box_len - box_width
			box_max = box_len - 1
			box = box_whole_string[box_min..box_max].to_i.to_s.rjust(box_width,"0")
			box_major_min = 0
			box_major_max = box_len - box_width - 1
			box_major = box_whole_string[box_major_min..box_major_max].to_i.to_s.rjust(box_major_width,"0")

			current_item = items_by_guid[record["GUID"]]
			if current_item.nil?
				pd.logMessage("Could not find item for GUID: #{record["GUID"]}")
			end

			# Setup placeholders
			resolver.clear
			resolver.setPath("export_directory",export_directory)

			# User "static" placeholders, evaluated earlier on so that
			# they should be able to contain other placeholders
			resolver.setPath("user_1",user_value_1)
			resolver.setPath("user_2",user_value_2)
			resolver.setPath("user_3",user_value_3)
			resolver.setPath("user_4",user_value_4)
			resolver.setPath("user_5",user_value_5)

			# Placeholders which are based on custom metadata
			if use_custom_placeholders
				item_custom_metadata = item.getCustomMetadata
				resolver.setPath("custom_1",item_custom_metadata[custom_field_1] || "")
				resolver.setPath("custom_2",item_custom_metadata[custom_field_2] || "")
				resolver.setPath("custom_3",item_custom_metadata[custom_field_3] || "")
				resolver.setPath("custom_4",item_custom_metadata[custom_field_4] || "")
				resolver.setPath("custom_5",item_custom_metadata[custom_field_5] || "")
			end

			# General placeholders
			resolver.set("box",box)
			resolver.set("box_major",box_major)
			current_item_localised_name = current_item.getLocalisedName.gsub(/[\\\/\n\r\t]/,"_")
			if current_item_localised_name =~ /\./
				resolver.set("name",File.basename(current_item_localised_name,File.extname(current_item_localised_name)))
			else
				resolver.set("name",current_item_localised_name)
			end
			resolver.set("fullname",current_item_localised_name)
			resolver.set("guid",current_item.getGuid)
			resolver.set("sub_guid",current_item.getGuid[0..2])
			resolver.set("md5",current_item.getDigests.getMd5 || "NO_MD5")
			resolver.set("type",current_item.getType.getLocalisedName)
			resolver.set("mime_type",current_item.getType.getName)
			resolver.set("kind",current_item.getType.getKind.getName)
			
			# Resolve extension placeholder
			extension = "BIN"
			if current_item.getKind.getName == "email"
				extension = natives_email_format
			else
				extension = current_item.getCorrectedExtension
				extension ||= current_item.getOriginalExtension
				extension ||= current_item.getType.getPreferredExtension
				extension ||= "BIN"
			end
			resolver.set("extension",extension)

			resolver.set("custodian",current_item.getCustodian || "NO_CUSTODIAN")
			resolver.set("evidence_name",current_item.getRoot.getLocalisedName)
			resolver.set("case_name",$current_case.getName)
			
			# Resolve an item date value for placeholder
			current_item_date = current_item.getDate
			if !current_item_date.nil?
				resolver.set("item_date",current_item_date.toString("YYYYMMdd"))
			else
				resolver.set("item_date","00000000")
			end
			if !current_item_date.nil?
				resolver.set("item_date_time",current_item_date.toString("YYYYMMdd-HHmmss"))
			else
				resolver.set("item_date_time","00000000-000000")
			end

			# Resolve an item path value that is able to be used in file system path
			current_item_path = current_item.getPath.to_a
			current_item_path.pop
			resolver.set("item_path",current_item_path.map{|i|i.getLocalisedName.gsub(/[\\\/\.\n\r\t]/,"_")}.join("\\"))

			# If we are exporting from a production set we can let the user make use of doc_id as we will
			# have data available to resolve this
			if enable_docid
				resolver.set("docid",docid_lookup[current_item])
			end

			# Placeholder which are based on top level item (or current item if above top level)
			top_level_item = current_item.getTopLevelItem || current_item
			resolver.set("top_level_guid",top_level_item.getGuid)
			resolver.set("top_level_sub_guid",top_level_item.getGuid[0..2])
			resolver.set("top_level_name",File.basename(top_level_item.getLocalisedName,File.extname(top_level_item.getLocalisedName)))
			resolver.set("top_level_fullname",top_level_item.getLocalisedName)
			top_level_item_path = top_level_item.getPath.to_a
			top_level_item_path.pop
			resolver.set("top_level_item_path",top_level_item_path.map{|i|i.getLocalisedName.gsub(/[\\\/\.\n\r\t]/,"_")}.join("\\"))

			filtered_path_names = []
			# Only calculate path names if we really need to since it could cause a performance hit
			if generate_filtered_path || (use_custom_placeholders && (resolver.get("custom_1") =~ fp_regex || resolver.get("custom_2") =~ fp_regex ||
				resolver.get("custom_3") =~ fp_regex || resolver.get("custom_4") =~ fp_regex || resolver.get("custom_5") =~ fp_regex))
				filtered_path_names = current_item_path
				filtered_path_names = filtered_path_names.reject{|path_item| filtered_path_mime_types[path_item.getType.getName]}
				filtered_path_names = filtered_path_names.map{|path_item|path_item.getLocalisedName}
				resolver.set("filtered_item_path",filtered_path_names.map{|n|n.gsub(/[\\\/\.\n\r\t]/,"_")}.join("\\"))
				if filter_dat_item_path && !record["Path Name"].nil?
					record["Path Name"] = "/"+filtered_path_names.join("/")
				end
			end

			# DEBUG - In case you need to take a peak at all the placeholder data for each item.
			# 	pd.logMessage("="*20)
			# 	resolver.getPlaceholderData.each do |k,v|
			# 	pd.logMessage("#{k} => #{v}")
			# end

			# I know these following method calls are ugly, it was a quick refactor vs a redesign into full on classes and such.  This could
			# probably be cleaned up at some point, but for now gets the job done.

			# Process text
			if export_text
				restructure_product(record,"TEXTPATH","Text",resolver,current_item,text_template,temp_export_directory,export_directory,placeholder_tags)
			end

			# Process natives
			if export_natives
				restructure_product(record,"ITEMPATH","Native",resolver,current_item,natives_template,temp_export_directory,export_directory,placeholder_tags)
			end

			# Process PDFs
			if export_pdf
				restructure_product(record,"PDFPATH","PDF",resolver,current_item,pdf_template,temp_export_directory,export_directory,placeholder_tags)
			end

			# Process TIFFs
			if export_tiff
				restructure_product(record,"TIFFPATH","TIFF",resolver,current_item,tiff_template,temp_export_directory,export_directory,placeholder_tags)
			end

			# Handle additional loadfile: CSV
			if export_csv
				# If were on first record, write headers first
				if record_number == 1
					csv << record.keys
				end
				# Write record values
				csv << record.values
			end

			if export_xlsx
				# If were on first record, write headers first
				if record_number == 1
					sheet << record.keys
				end
				# Write record values
				sheet << record.values
			end
		end

		# If we're exporting additional load file formats (CSV/XLSX) we need to close them out
		if export_csv
			csv.close
		end

		if export_xlsx
			sheet.auto_fit_columns
			xlsx.save(xlsx_file)
			xlsx.dispose
		end

		# Cleanup what remains of the initial temporary export location.  Shouldn't be much as we should have moved
		# everything into the new structure.
		pd.setMainStatusAndLogIt("Deleting temporary export directory...")
		org.apache.commons.io.FileUtils.deleteDirectory(java.io.File.new(temp_export_directory))

		# Put the progress dialog into a completed state
		pd.setCompleted
	end
end