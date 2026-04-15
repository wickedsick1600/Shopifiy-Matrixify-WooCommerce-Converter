/**
 * GlossyLounge: Post-import variation fixer (all-in-one).
 *
 * Run ONCE via Code Snippets after importing products.
 * 1. Splits combined pa_size terms (XS|S|M|L|XL → individual terms)
 * 2. Cleans ="..." from term names
 * 3. Sets product type to "variable" for all products with variations
 * 4. Forces pa_size attribute with "Used for variations" checked
 * 5. Fixes variation meta keys (attribute_size → attribute_pa_size)
 *
 * Remove this snippet after running.
 */

if ( ! defined( 'ABSPATH' ) ) exit;

global $wpdb;
set_time_limit( 600 );

$taxonomy = 'pa_size';
$stats = [ 'terms_split' => 0, 'terms_cleaned' => 0, 'products_fixed' => 0, 'variations_fixed' => 0 ];

$name_map = [
    'xs' => 'XS', 's' => 'S', 'm' => 'M', 'l' => 'L',
    'xl' => 'XL', '2xl' => '2XL', 'xxl' => '2XL', 'one-size' => 'One Size',
];

// ── Helper: get or create a single pa_size term, return term_taxonomy_id ──
function glossy_ensure_term( $slug, $label = '' ) {
    global $wpdb, $name_map;
    $slug = sanitize_title( $slug );
    if ( ! $slug ) return null;

    $row = $wpdb->get_row( $wpdb->prepare(
        "SELECT tt.term_taxonomy_id FROM {$wpdb->terms} t
         INNER JOIN {$wpdb->term_taxonomy} tt ON t.term_id = tt.term_id
         WHERE t.slug = %s AND tt.taxonomy = 'pa_size' LIMIT 1", $slug
    ) );
    if ( $row ) return (int) $row->term_taxonomy_id;

    $nice = $label ?: ( isset( $name_map[ $slug ] ) ? $name_map[ $slug ] : strtoupper( $slug ) );
    $wpdb->insert( $wpdb->terms, [ 'name' => $nice, 'slug' => $slug, 'term_group' => 0 ] );
    $tid = $wpdb->insert_id;
    if ( ! $tid ) return null;
    $wpdb->insert( $wpdb->term_taxonomy, [
        'term_id' => $tid, 'taxonomy' => 'pa_size', 'description' => '', 'parent' => 0, 'count' => 0,
    ] );
    return (int) $wpdb->insert_id;
}

// ══════════════════════════════════════════════════════════════════════════
// STEP 1: Split all combined pa_size terms (names containing "|")
// ══════════════════════════════════════════════════════════════════════════

$combined = $wpdb->get_results(
    "SELECT t.term_id, t.name, t.slug, tt.term_taxonomy_id
     FROM {$wpdb->terms} t
     INNER JOIN {$wpdb->term_taxonomy} tt ON t.term_id = tt.term_id
     WHERE tt.taxonomy = 'pa_size' AND t.name LIKE '%|%'"
);

foreach ( $combined as $bad ) {
    $product_ids = $wpdb->get_col( $wpdb->prepare(
        "SELECT object_id FROM {$wpdb->term_relationships} WHERE term_taxonomy_id = %d",
        $bad->term_taxonomy_id
    ) );

    $parts = array_filter( array_map( 'trim', explode( '|', $bad->name ) ) );
    $individual_tt_ids = [];
    foreach ( $parts as $part ) {
        $clean = trim( str_replace( [ '="', '"' ], '', $part ) );
        $tt_id = glossy_ensure_term( sanitize_title( $clean ), $clean );
        if ( $tt_id ) $individual_tt_ids[] = $tt_id;
    }

    foreach ( $product_ids as $pid ) {
        $wpdb->delete( $wpdb->term_relationships, [
            'object_id' => (int) $pid, 'term_taxonomy_id' => (int) $bad->term_taxonomy_id,
        ] );
        foreach ( $individual_tt_ids as $tt_id ) {
            $wpdb->replace( $wpdb->term_relationships, [
                'object_id' => (int) $pid, 'term_taxonomy_id' => $tt_id, 'term_order' => 0,
            ] );
        }
    }

    $wpdb->delete( $wpdb->term_relationships, [ 'term_taxonomy_id' => (int) $bad->term_taxonomy_id ] );
    $wpdb->delete( $wpdb->term_taxonomy, [ 'term_taxonomy_id' => (int) $bad->term_taxonomy_id ] );
    $wpdb->delete( $wpdb->terms, [ 'term_id' => (int) $bad->term_id ] );
    $stats['terms_split']++;
}

// ══════════════════════════════════════════════════════════════════════════
// STEP 2: Clean ="..." from any term display names
// ══════════════════════════════════════════════════════════════════════════

$dirty = $wpdb->get_results(
    "SELECT t.term_id, t.name FROM {$wpdb->terms} t
     INNER JOIN {$wpdb->term_taxonomy} tt ON t.term_id = tt.term_id
     WHERE tt.taxonomy = 'pa_size' AND (t.name LIKE '%=%' OR t.name LIKE '%\"%')"
);

foreach ( $dirty as $d ) {
    $clean_name = trim( str_replace( [ '="', '"', "'" ], '', $d->name ) );
    $slug = sanitize_title( $clean_name );
    if ( isset( $name_map[ $slug ] ) ) $clean_name = $name_map[ $slug ];
    $wpdb->update( $wpdb->terms, [ 'name' => $clean_name ], [ 'term_id' => (int) $d->term_id ] );
    $stats['terms_cleaned']++;
}

// ══════════════════════════════════════════════════════════════════════════
// STEP 3: Fix every variable product
// ══════════════════════════════════════════════════════════════════════════

$variable_term_tt = $wpdb->get_var(
    "SELECT tt.term_taxonomy_id FROM {$wpdb->terms} t
     INNER JOIN {$wpdb->term_taxonomy} tt ON t.term_id = tt.term_id
     WHERE tt.taxonomy = 'product_type' AND t.slug = 'variable' LIMIT 1"
);

$parent_ids = $wpdb->get_col(
    "SELECT DISTINCT p.post_parent
     FROM {$wpdb->posts} p
     INNER JOIN {$wpdb->posts} parent
       ON parent.ID = p.post_parent AND parent.post_type = 'product'
       AND parent.post_status IN ('publish','draft','private')
     WHERE p.post_type = 'product_variation' AND p.post_parent > 0"
);

foreach ( $parent_ids as $pid ) {
    $pid = (int) $pid;

    // 3a. Set product type to "variable"
    if ( $variable_term_tt ) {
        $wpdb->query( $wpdb->prepare(
            "DELETE FROM {$wpdb->term_relationships}
             WHERE object_id = %d AND term_taxonomy_id IN
               (SELECT term_taxonomy_id FROM {$wpdb->term_taxonomy} WHERE taxonomy = 'product_type')",
            $pid
        ) );
        $wpdb->replace( $wpdb->term_relationships, [
            'object_id' => $pid, 'term_taxonomy_id' => (int) $variable_term_tt, 'term_order' => 0,
        ] );
    }

    // 3b. Force _product_attributes: pa_size with is_variation=1, is_taxonomy=1
    $attrs = get_post_meta( $pid, '_product_attributes', true );
    if ( ! is_array( $attrs ) ) $attrs = [];

    $new_attrs = [];
    foreach ( $attrs as $key => $data ) {
        if ( ! is_array( $data ) ) continue;
        $kl = strtolower( $key );
        if ( $kl === 'size' || ( strpos( $kl, 'size' ) !== false && $kl !== 'pa_size' ) ) continue;
        if ( $kl === 'pa_size' ) continue;
        $new_attrs[ $key ] = $data;
    }
    $new_attrs['pa_size'] = [
        'name' => 'pa_size', 'value' => '', 'position' => 0,
        'is_visible' => 1, 'is_variation' => 1, 'is_taxonomy' => 1,
    ];
    update_post_meta( $pid, '_product_attributes', $new_attrs );

    // 3c. Fix variation meta: rename attribute_size → attribute_pa_size
    $wpdb->query( $wpdb->prepare(
        "UPDATE {$wpdb->postmeta} SET meta_key = 'attribute_pa_size'
         WHERE meta_key = 'attribute_size'
         AND post_id IN (
           SELECT ID FROM {$wpdb->posts}
           WHERE post_type = 'product_variation' AND post_parent = %d
         )", $pid
    ) );

    // 3d. Ensure every variation has attribute_pa_size set
    $variations = $wpdb->get_results( $wpdb->prepare(
        "SELECT p.ID as vid, pm.meta_value as size_slug
         FROM {$wpdb->posts} p
         LEFT JOIN {$wpdb->postmeta} pm
           ON pm.post_id = p.ID AND pm.meta_key = 'attribute_pa_size'
         WHERE p.post_type = 'product_variation' AND p.post_parent = %d
         ORDER BY p.ID ASC", $pid
    ) );

    $product_sizes = $wpdb->get_col( $wpdb->prepare(
        "SELECT t.slug FROM {$wpdb->terms} t
         INNER JOIN {$wpdb->term_taxonomy} tt ON t.term_id = tt.term_id
         INNER JOIN {$wpdb->term_relationships} tr ON tt.term_taxonomy_id = tr.term_taxonomy_id
         WHERE tr.object_id = %d AND tt.taxonomy = 'pa_size'", $pid
    ) );

    foreach ( $variations as $i => $v ) {
        $vid = (int) $v->vid;
        $slug = $v->size_slug;

        if ( ! $slug && ! empty( $product_sizes ) ) {
            $slug = $product_sizes[ $i % count( $product_sizes ) ];
        }
        if ( ! $slug ) continue;

        $existing = $wpdb->get_var( $wpdb->prepare(
            "SELECT meta_value FROM {$wpdb->postmeta}
             WHERE post_id = %d AND meta_key = 'attribute_pa_size'", $vid
        ) );

        if ( $existing === null ) {
            $wpdb->insert( $wpdb->postmeta, [
                'post_id' => $vid, 'meta_key' => 'attribute_pa_size', 'meta_value' => $slug,
            ] );
        } elseif ( $existing !== $slug && ! $existing ) {
            $wpdb->update( $wpdb->postmeta,
                [ 'meta_value' => $slug ],
                [ 'post_id' => $vid, 'meta_key' => 'attribute_pa_size' ]
            );
        }

        $stats['variations_fixed']++;
    }

    delete_transient( 'wc_product_children_' . $pid );
    delete_transient( 'wc_var_prices_' . $pid );
    wp_cache_delete( $pid, 'posts' );
    $stats['products_fixed']++;
}

// Update term counts
$wpdb->query(
    "UPDATE {$wpdb->term_taxonomy} tt SET count = (
        SELECT COUNT(*) FROM {$wpdb->term_relationships} tr
        WHERE tr.term_taxonomy_id = tt.term_taxonomy_id
     ) WHERE tt.taxonomy = 'pa_size'"
);

wp_die( sprintf(
    'Done.<br>Combined terms split: %d<br>Term names cleaned: %d<br>Products fixed: %d<br>Variations fixed: %d<br><br>Remove this snippet now.',
    $stats['terms_split'], $stats['terms_cleaned'], $stats['products_fixed'], $stats['variations_fixed']
) );