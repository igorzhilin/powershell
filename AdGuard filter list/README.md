# Generate AdGuard filter list with powershell
In this example, I create a filter list to block Huawei domains with ad blocking app [AdGuard](https://adguard.com/).
The filter list should also work with ad blocking browser extension [uBlock Origin](https://ublockorigin.com/).

The `.hosts` files are sourced from here: [https://codeberg.org/Cyanic76/Hosts](https://codeberg.org/Cyanic76/Hosts)

## Input
* URL(s) with `.hosts` files containing the domains to block

## Output
* URL of the plain text document containing the domains extracted from the `.hosts` files 👉 take this URL and add it to AdGuard

## Add the generated filter list to AdGuard
* Settings ➡ Content blocking ➡ Filters ➡ Custom filters ➡ ➕ NEW CUSTOM FILTER ➡ Add the URL ➡ IMPORT

## Example outputs
* Huawei filter list for AdGuard: [http://sprunge.us/bZNyaw](http://sprunge.us/bZNyaw)
* Complete filter list for AdGuard based on full **[Cyanic's Hosts list](https://hosts.cyanic.me/cyanicHosts.txt)** (contains `localhost` and `0.0.0.0` for some reason, but otherwise looks ok): [http://sprunge.us/7iPkRg](http://sprunge.us/7iPkRg)

## Screenshot
![](https://i.imgur.com/YbZHK4r.png)