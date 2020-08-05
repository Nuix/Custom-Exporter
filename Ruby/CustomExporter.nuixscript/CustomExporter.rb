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
load File.join(script_directory,"DAT.rb")

# Load class for working with OPT file
load File.join(script_directory,"OPT.rb")

# Load class for exporting XLSX
load File.join(script_directory,"Xlsx.rb")

# Require CSV library
require 'csv'

java_import org.apache.commons.io.FilenameUtils

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
dialog.setHelpUrl("https://github.com/Nuix/Custom-Exporter")

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

built_in_headers = [
	"DOCID",
	"PARENT_DOCID",
	"ATTACH_DOCID",
	"BEGINBATES",
	"ENDBATES",
	"BEGINGROUP",
	"ENDGROUP",
	"PAGECOUNT",
	"ITEMPATH",
	"TEXTPATH",
	"PDFPATH",
	"TIFFPATH",
]
headers_tab = dialog.addTab("headers_tab","Nuix DAT Headers")
built_in_headers.each do |built_in_header|
	headers_tab.appendTextField("rename_#{built_in_header.downcase}","Rename '#{built_in_header}' to",built_in_header)
end

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
tiff_tab.appendRadioButton("multi_page_tiff","Multi-page","tiff_type_group",true)
tiff_tab.appendRadioButton("single_page_tiff","Single-page","tiff_type_group",false)
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
native_tab.appendComboBox("natives_email_format","Email Format",["msg","eml","html","mime_html","dxl"])
native_tab.appendCheckBox("include_attachments","Include Attachments on Emails",true)
native_tab.enabledOnlyWhenChecked("natives_template","export_natives")
native_tab.enabledOnlyWhenChecked("natives_email_format","export_natives")

# Placeholder settings tab
placeholders_tab = dialog.addTab("placeholders_tab","Placeholders")

placeholders_tab.appendSpinner("max_item_path_segment_length","Maximum {item_path} segment length",0,0,10000000,1)

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
worker_tab.appendCheckBox("delete_temp_directory","Delete Temp Directory on Completion",true)

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
	if values["export_tiff"] 
		if (values["tiff_template"].nil? || values["tiff_template"].strip.empty?)
			CommonDialogs.showError("TIFF Path Template cannot be empty.")
			next false
		end

		multi_page_tiff = values["multi_page_tiff"] == true
		if !multi_page_tiff && values["tiff_template"] !~ /\{page\}/ && values["tiff_template"] !~ /\{page_4\}/
			message = "You have specified export of single page TIFFs, but have not provided a page\n"
			message += "format placeholder in your TIFF pathing template.  Please include either of\n"
			message += "the following placeholders in your TIFF path template:\n"
			message += "{page} - Page number without zero fill\n"
			message += "{page_4} - Page number zero filled 4 characters wide"
			CommonDialogs.showError(message,"No Page Placeholder in TIFF Path Template")
			next false
		end
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
def restructure_product(record,product_path_field,product_name,resolver,current_item,template,temp_export_directory,export_directory,placeholder_tags,pd)
	xref = {}
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

		# Handle things like:
		# - Too long file path
		# - Suffix when name collides with file already on disk
		pathing_data.each do |path_entry|
			updated_fullpath = path_entry[:updated_fullpath] # Absolute path
			updated_path = path_entry[:updated_path] # Relative path for loadfile

			# We need to handle path too long here since FileUtils.copyFile doesn't seem to be
			# throwing exceptions for this
			was_truncated = false
			original_updated_fullpath = updated_fullpath
			if updated_fullpath.size >= 260
				pd.logMessage("Resolved path equals/exceeds 260 characters and will be truncated for item with GUID #{current_item.getGuid}")
				resolved_file_name = java.io.File.new(updated_fullpath).getName
				updated_fullpath = File.join(export_directory,"PATH_TOO_LONG",resolved_file_name)
				was_truncated = true
			end

			copy = 1
			temp_fullpath = updated_fullpath
			while java.io.File.new(temp_fullpath).exists
				extension = FilenameUtils.getExtension(updated_fullpath)
				file_base_name = FilenameUtils.getBaseName(updated_fullpath)
				directory = java.io.File.new(updated_fullpath).getParentFile.getAbsolutePath
				temp_fullpath = File.join(directory,"#{file_base_name}_#{copy}.#{extension}")
				copy += 1
			end
			updated_fullpath = temp_fullpath

			updated_fullpath = updated_fullpath.gsub("/","\\")
			updated_path = updated_fullpath.gsub(/^#{Regexp.escape(export_directory)}\\/i,".\\")

			if was_truncated
				File.open(File.join(export_directory,"PATH_TOO_LONG.TXT"),"a:utf-8") do |path_log|
					path_log.puts("="*20)
					path_log.puts("GUID: #{current_item.getGuid}")
					path_log.puts("Product: #{product_name}")
					path_log.puts("Resolved path of #{original_updated_fullpath.size} characters equals/exceeds 260 character limit")
					path_log.puts(" Resolved Path: #{original_updated_fullpath}")
					path_log.puts("Truncated Path: #{updated_fullpath}")
				end
			end

			path_entry[:updated_fullpath] = updated_fullpath
			path_entry[:updated_path] = updated_path
		end
		
		had_error = false
		pathing_data.each_with_index do |path_entry,path_entry_index|
			updated_fullpath = path_entry[:updated_fullpath]
			updated_path = path_entry[:updated_path]

			xref[original_path] = updated_path

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
	
	return xref
end

def find_additional_pages(temp_export_directory,first_page_file)
	result = []
	first_page_file = first_page_file.gsub(/^\.\\/,"")
	dir = File.dirname(first_page_file)
	ext = File.extname(first_page_file)
	name = File.basename(first_page_file, ext)
	index = 1
	while true
		maybe_exists = File.join(temp_export_directory,dir,"#{name}_#{index}#{ext}")
		puts "Checking for existence of: #{maybe_exists}"
		if java.io.File.new(maybe_exists).exists
			result << File.join(dir,"#{name}_#{index}#{ext}").gsub(/\//,"\\")
			index += 1
			puts "Found additional page image: #{maybe_exists}"
		else
			break
		end
	end
	return result
end

def coalesce(input,fallback="NO_VALUE")
	if input.nil? || input.strip.empty?
		return fallback
	else
		return input
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

		header_renames = {}
		built_in_headers.each do |built_in_header|
			header_renames[built_in_header] = values["rename_#{built_in_header.downcase}"]
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
		multi_page_tiff = values["multi_page_tiff"] == true

		export_directory = values["export_directory"]
		temp_export_directory = File.join(export_directory,"TEMP")
		max_item_path_segment_length = values["max_item_path_segment_length"]
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

		delete_temp_directory = values["delete_temp_directory"]
		
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
				:multiPageTiff => multi_page_tiff,
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
		exported_temp_opt = File.join(temp_export_directory,"loadfile.opt")
		final_opt = File.join(export_directory,"loadfile.opt")
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

		tiff_xref = {}

		# Specify header renames
		DAT.when_modify_headers do |headers|
			modified_headers = []
			headers.each do |header|
				new_header = header_renames[header]
				if new_header.nil?
					modified_headers << header
				else
					modified_headers << new_header
				end
			end
			next modified_headers
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
				cm = current_item.getCustomMetadata
				# Coalesce will make sure nil or empty values get a value
				# of "NO_VALUE" since an empty value could break path strings
				cmv1 = coalesce("#{cm[custom_field_1]}")
				cmv2 = coalesce("#{cm[custom_field_2]}")
				cmv3 = coalesce("#{cm[custom_field_3]}")
				cmv4 = coalesce("#{cm[custom_field_4]}")
				cmv5 = coalesce("#{cm[custom_field_5]}")

				# Furthermore, we need to remove illegal path characters from these
				# values while still allowing path separators to get through in case
				# someone is using the field to specify varying pathing
				cmv1 = PlaceholderResolver.cleanPathString(cmv1)
				cmv2 = PlaceholderResolver.cleanPathString(cmv2)
				cmv3 = PlaceholderResolver.cleanPathString(cmv3)
				cmv4 = PlaceholderResolver.cleanPathString(cmv4)
				cmv5 = PlaceholderResolver.cleanPathString(cmv5)

				resolver.setPath("custom_1",cmv1)
				resolver.setPath("custom_2",cmv2)
				resolver.setPath("custom_3",cmv3)
				resolver.setPath("custom_4",cmv4)
				resolver.setPath("custom_5",cmv5)
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
				if natives_email_format == "mime_html"
					extension = "mht"
				else
					extension = natives_email_format
				end
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
			current_item_path = current_item.getLocalisedPathNames.to_a
			current_item_path.pop
			current_item_path = current_item_path.map do |localised_path_name|
				# Replace some characters with underscores
				segment = localised_path_name.gsub(/[\\\/\.\n\r\t]+/,"_")

				# If user specified max segment length, enforce that here
				if max_item_path_segment_length > 0 && segment.size > max_item_path_segment_length
					segment = segment[0..max_item_path_segment_length-1]
				end

				# Its possible for leading or trailing whitespace characters at this
				# point which wont work so we need to strip them off
				segment = segment.strip

				# Finally, there is a small chance the somehow we have a segment of length 0
				# so we will just make sure something gets through for this segment
				if segment.size == 0
					segment = "_"
				end

				next segment
			end
			resolver.set("item_path",current_item_path.join("\\"))

			# If we are exporting from a production set we can let the user make use of doc_id as we will
			# have data available to resolve this
			if enable_docid
				docid_to_use = docid_lookup[current_item]
				if docid_to_use.nil? || docid_to_use.strip.empty?
					resolver.set("docid",current_item.getGuid)
				else
					resolver.set("docid",docid_to_use)
				end
			else
				resolver.set("docid",current_item.getGuid)
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

			# These will be replaced with blank values unless we are exporting single page
			# images (TIFF/JPG).  When exporting single page images, will be updated per page
			# as each page is being restructured.
			resolver.set("page","")
			resolver.set("page_4","")

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
				restructure_product(record,"TEXTPATH","Text",resolver,current_item,text_template,temp_export_directory,export_directory,placeholder_tags,pd)
			end

			# Process natives
			if export_natives
				restructure_product(record,"ITEMPATH","Native",resolver,current_item,natives_template,temp_export_directory,export_directory,placeholder_tags,pd)
			end

			# Process PDFs
			if export_pdf
				restructure_product(record,"PDFPATH","PDF",resolver,current_item,pdf_template,temp_export_directory,export_directory,placeholder_tags,pd)
			end

			# Process TIFFs
			if export_tiff
				if multi_page_tiff
					xref = restructure_product(record,"TIFFPATH","TIFF",resolver,current_item,tiff_template,temp_export_directory,export_directory,placeholder_tags,pd)
					tiff_xref.merge!(xref)
				else
					base_tiff_path = record["TIFFPATH"]
					page = 1

					# Set page placeholders
					resolver.set("page","#{page}")
					resolver.set("page_4",page.to_s.rjust(4,"0"))

					xref = restructure_product(record,"TIFFPATH","TIFF",resolver,current_item,tiff_template,temp_export_directory,export_directory,placeholder_tags,pd)
					tiff_xref.merge!(xref)

					find_additional_pages(temp_export_directory,base_tiff_path).each do |additional_page_file|
						page += 1

						# Set page placeholders
						resolver.set("page","#{page}")
						resolver.set("page_4",page.to_s.rjust(4,"0"))

						# restructure_product method was built to handle restructuring item level files, not
						# individual pages of an item.  Part of this expectation is that it gets the current item's
						# DAT record as a hash.  To allow for restructuring of individual pages, we provide the item
						# level record, but stuff the TIFFPATH field with the individual page image's file path each
						# time we restructure a page level image.
						record["TIFFPATH"] = additional_page_file

						xref = restructure_product(record,"TIFFPATH","TIFF",resolver,current_item,tiff_template,temp_export_directory,export_directory,placeholder_tags,pd)
						tiff_xref.merge!(xref)
					end
				end
			end

			# Handle additional loadfile: CSV
			if export_csv
				# If were on first record, write headers first
				if record_number == 1
					csv << DAT.modify_headers(record.keys)
				end
				# Write record values
				csv << record.values
			end

			if export_xlsx
				# If were on first record, write headers first
				if record_number == 1
					sheet << DAT.modify_headers(record.keys)
				end
				# Write record values
				sheet << record.values
			end
		end

		# Fix up OPT if we're exporting TIFFs
		if export_tiff
			puts tiff_xref.inspect
			OPT.transpose_each(exported_temp_opt,final_opt) do |opt_record|
				path = tiff_xref[opt_record.path]
				if !path.nil?
					opt_record.path = path.gsub(/^\.\\/,"")

					# When exporting a production set, the OPT produced by Nuix during the temp export
					# phase will contain DOCIDs for ID in each OPT record.  When not exporting from a production
					# set we get vague IDs like: 001.001.002
					# Here we add logic to instead use the resolved image's file name as the OPT ID.
					if !values["use_production_set"]
						ext = File.extname(path)
						name = File.basename(path, ext)
						opt_record.id = name
					end
				end
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

		if delete_temp_directory
			# Cleanup what remains of the initial temporary export location.  Shouldn't be much as we should have moved
			# everything into the new structure.
			pd.setMainStatusAndLogIt("Deleting temporary export directory...")
			org.apache.commons.io.FileUtils.deleteDirectory(java.io.File.new(temp_export_directory))
		else
			pd.setMainStatusAndLogIt("Not deleting temporary export directory, as per settings.")
		end

		# Put the progress dialog into a completed state
		pd.setCompleted
	end
end