CREATE SCHEMA IF NOT EXISTS example;

CREATE  TABLE IF NOT EXISTS example.products(
    id bigserial not null,
    price numeric,
    product_name varchar(255),
    stock int,
    credential uuid not null
);