# Smart MegaMenu for SPO

## Add Url to Terms

Either add localCustomProperties to term by the name `url` (must lower case) OR change the TermSet to be Navigation type and fill `SimpleLinkUrl`.

DOES NOT WORK WITH FRIENDLY URLS

## Settings List

### Automativally Created

Your MUST
* create a termset
* add the guid to the `TermSetGuid` row @`Value` column

#### other settings
* You may choose between `Simple`, `Blue` or `EndlessExpand` in the `MegaType` (TODO...)
* You may set the `Direction` row to `rtl` or `ltr` values to override the `direction` css value

![How List Looks](https://github.com/bresleveloper/SPFx-Smart_MegaMenu/blob/master/readmeImage.png?raw=true)


## RoadMap
* `Blue` as in mockup
* `Simple` as in FF
* `EndlessExpand` as in OOTB teamsite
