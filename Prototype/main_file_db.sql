CREATE DATABASE database_tdb;
USE database_tdb;

BACKUP DATABASE database_tdb
TO DISK = 'file_paths'
WITH DIFFERENTIAL;

show tables;
select * from pt_info;

CREATE TABLE pt_info (
	pt_id int NOT NULL PRIMARY KEY AUTO_INCREMENT,
    pt_name varchar(255),
    pt_name_short varchar(255),
    pt_alamat varchar(255),
    phone varchar(255),
    NPWP varchar(255)
);

CREATE TABLE product (
	product_id int NOT NULL PRIMARY KEY auto_increment,
    material_no varchar(255),
    description varchar (255),
    unit_of_measurement varchar (255),
    quantity int
);

CREATE TABLE product_purchase_order (
	purchase_no bigint NOT NULL,
    product_id int NOT NULL,
    qty int,
    disc int,
    unit_price bigint,
    currency varchar(255),
    total bigint,
    remark varchar (255),
    FOREIGN KEY (product_id) REFERENCES product(product_id)
);

CREATE TABLE site_info (
	site_id int NOT NULL PRIMARY KEY AUTO_INCREMENT,
    pt_id int,
    site_name varchar(255),
    site_alamat varchar(255),
    FOREIGN KEY (pt_id) REFERENCES pt_info(pt_id)
);

CREATE TABLE purchase_order (
	purchase_no bigint NOT NULL PRIMARY KEY,
    purchase_date date NOT NULL,
    pt_id int NOT NULL,
    site_id int NOT NULL,
    product_qty int,
    FOREIGN KEY (pt_id) REFERENCES pt_info(pt_id),
    FOREIGN KEY (site_id) REFERENCES site_info(site_id)
);

CREATE TABLE delivery_order (
	delivery_id varchar(255) NOT NULL PRIMARY KEY,
    purchase_no bigint NOT NULL,
    delivery_date date NOT NULL,
    invoice_status bool default 0,
    invoice_no varchar(255),
    note varchar (255),
    revision int,
	FOREIGN KEY (purchase_no) REFERENCES purchase_order(purchase_no),
    FOREIGN KEY (invoice_no) REFERENCES invoice(invoice_no)
);

CREATE TABLE invoice (
	invoice_no int NOT NULL PRIMARY KEY,
    invoice_date date,
    bill_to_pt_id int NOT NULL,
    ship_to_site_id int NOT NULL,
    term_of_payment date,
    shipping_via varchar (255),
    ship_date date,
    purchase_no bigint NOT NULL,
    description varchar (255),
    approve bool DEFAULT 0,
    FOREIGN KEY (bill_to_pt_id) REFERENCES pt_info(pt_id),
    FOREIGN KEY (ship_to_site_id) REFERENCES site_info(site_id),
    FOREIGN KEY (purchase_no) REFERENCES purchase_order(purchase_no)
);

USE database_tdb;

drop table site_info;
drop table pt_info;
drop table purchase_order;
drop table product;
drop table delivery_order_no;
drop table delivery_order;
drop table invoice;
drop table product_purchase_order;

DELETE FROM site_info;
DELETE FROM pt_info;
DELETE FROM purchase_order;
DELETE FROM product;
DELETE FROM delivery_order_no;
DELETE FROM delivery_order;
DELETE FROM invoice;
DELETE FROM product_purchase_order;
