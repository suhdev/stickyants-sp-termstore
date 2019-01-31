export interface IValue {
    id:string|number;
    label:string; 
}

export interface ITermStore {
    createTerm(parentTerm:SP.Taxonomy.Term, name:string, locale:number, guid:SP.Guid, properties?:IValue[]):Promise<SP.Taxonomy.Term>; 
    getParentTermById(termId:string):Promise<SP.Taxonomy.Term>; 
    getTerms(termIds:string[]):Promise<SP.Taxonomy.Term[]>;
    getTermsByTermId(termId:string):Promise<SP.Taxonomy.Term[]>;
    getTermsByTermSetId(termSetId:string):Promise<SP.Taxonomy.Term[]>;  
    getAllTermsByTermSetId(termSetId: string):Promise<SP.Taxonomy.Term[]>;
    termSetIdFromTaxonomyField(fieldInternalName:string):Promise<string>;
    getTermsByIds(ids:string[]):Promise<SP.Taxonomy.TermCollection>; 
    getParentTermByTerm(term:SP.Taxonomy.Term):Promise<SP.Taxonomy.Term>; 
    getTopLevelParentOfTerm(id:string):Promise<SP.Taxonomy.Term>;
    getTermParents(termId:string):Promise<SP.Taxonomy.Term[]>;
    getParentThatSatisfies(id:string,fn:(v:SP.Taxonomy.Term)=>boolean):Promise<SP.Taxonomy.Term>;
    getTermLabelsById(termId:string):Promise<SP.Taxonomy.Label[]>;
    getLabelsForTerms(terms:SP.Taxonomy.Term[]):Promise<SP.Taxonomy.LabelCollection[]>; 
    getTermsSubTreeFlat(termId:string,list?:SP.Taxonomy.Term[]):Promise<SP.Taxonomy.Term[]>;
    getAllTermSetsInSiteCollectionGroup(createIfMissing:boolean):Promise<SP.Taxonomy.TermSet[]>;
    getTermLabels(term:SP.Taxonomy.Term):Promise<SP.Taxonomy.Label[]>;
    getTermById(id:string):Promise<SP.Taxonomy.Term>;
    getSiteCollectionTermGroup(createIfMissing):Promise<SP.Taxonomy.TermGroup>;
}

export type TermStoreContextCallback<T> = (ctx: SP.ClientContext, session: SP.Taxonomy.TaxonomySession, tstore: SP.Taxonomy.TermStore, execute: (result: T) => void) => void; 
export function createTermStore():ITermStore{
    var cache = {
        termSetByTermId:{}, 
        termsByTermSetId:{},
        parentsByTermId:{}
    };

    function createExecutionContext<T>(fn: TermStoreContextCallback<T>,siteUrl?:string){
        let ctx = new SP.ClientContext(siteUrl||_spPageContextInfo.siteAbsoluteUrl); 
        let session = SP.Taxonomy.TaxonomySession.getTaxonomySession(ctx); 
        let tstore = session.getDefaultSiteCollectionTermStore(); 
        return new Promise<T>((resolve,reject)=>{
            fn(ctx, session, tstore, function (result:any) {
                ctx.executeQueryAsync(()=>{
                    resolve(result); 
                }, reject);
            }); 
        });
    }

    function getTermsByTermSetId(termSetId:string):Promise<SP.Taxonomy.Term[]>{
        return createExecutionContext<SP.Taxonomy.TermCollection>((ctx,session,store,execute)=>{
            let tset = store.getTermSet(new SP.Guid(termSetId));
            let terms = tset.get_terms();
            ctx.load(terms); 
            execute(terms);
        })
        .then((terms:SP.Taxonomy.TermCollection)=>{
            return terms.get_data(); 
        }); 
    }

    function getTermsByTermId(termId:string):Promise<SP.Taxonomy.Term[]>{
        return createExecutionContext<SP.Taxonomy.TermCollection>((ctx, session, store, execute) => {
            let tset = store.getTerm(new SP.Guid(termId));
            let terms = tset.get_terms();
            ctx.load(terms);
            execute(terms);
        })
        .then((terms: SP.Taxonomy.TermCollection) => {
            return terms.get_data();
        }); 
    }

    function getTermParentPaths(path:string){
        var paths = path.split(';');
        paths.pop(); 
        var prev = ''; 
        var pp = []; 
        for(var p of paths){
            prev = prev?prev +';'+p:p; 
            pp.push(prev); 
        }
        return pp; 
    }

    async function getTermParents(termId:string):Promise<SP.Taxonomy.Term[]>{
        var termPaths:string[] = []; 
        if (cache.parentsByTermId[termId]){
            return cache.parentsByTermId[termId]; 
        }
        return createExecutionContext<SP.Taxonomy.Term[]>(async (ctx, session, store, execute) => {
            let term = store.getTerm(new SP.Guid(termId));
            let tset = term.get_termSet(); 
            let terms = tset.getAllTerms(); 
            ctx.load(terms); 
            ctx.load(term); 
            await executeOnContext(ctx);
            termPaths = getTermParentPaths(term.get_pathOfTerm()); 
            var parents = terms.get_data().filter(e=>{
                return _.some(termPaths,v=>v === e.get_pathOfTerm()); 
            }); 
            execute(cache.parentsByTermId[termId] = parents);
        });
    }

    function getTermsByTermIds(...termIds:string[]):Promise<SP.Taxonomy.Term[][]>{
        var taxTerms:SP.Taxonomy.TermCollection[] = []; 
        return createExecutionContext<SP.Taxonomy.TermCollection>((ctx, session, store, execute) => {
            let terms:SP.Taxonomy.TermCollection; 
            for(var termId of termIds){
                let tset = store.getTerm(new SP.Guid(termId));
                terms = tset.get_terms();
                taxTerms.push(terms); 
                ctx.load(terms);
            }
            execute(terms);
        })
        .then((terms: SP.Taxonomy.TermCollection) => {
            return taxTerms.map(e=>e.get_data());
        }); 
    }

    function getTermSetByTermId(termId:string){
        return createExecutionContext<SP.Taxonomy.TermSet>((ctx, session, store, execute) => {
            var term = store.getTerm(new SP.Guid(termId)); 
            var tset = term.get_termSet(); 
            ctx.load(tset); 
            execute(tset);
        }); 
    }

    async function getTermsSubTreeFlat(termId:string,list:SP.Taxonomy.Term[] = []):Promise<SP.Taxonomy.Term[]>{
        var isTermSetCached = cache.termSetByTermId[termId]?true:false; 
        let termset:SP.Taxonomy.TermSet = cache.termSetByTermId[termId] || await getTermSetByTermId(termId); 
        cache.termSetByTermId[termId] = termset; 
        var isTermsCached = cache.termsByTermSetId[termset.get_id().toString()]?true:false; 
        let terms:SP.Taxonomy.Term[] = cache.termsByTermSetId[termset.get_id().toString()] || await new Promise((res,rej)=>{
            var ctx = termset.get_context();
            var terms = termset.getAllTerms();
            ctx.load(terms);
            ctx.executeQueryAsync(()=>{
                res(terms.get_data()); 
            },(c,err)=>{
                rej(err); 
            })
        }); 
        cache.termsByTermSetId[termset.get_id().toString()] = terms; 
        if (!isTermsCached){
            for(var t of terms){
                cache.termSetByTermId[t.get_id().toString()] = termset; 
            }
        }
        var parentTerm = _.find(terms,e=>e.get_id().toString() === termId); 
        if (parentTerm){
            return terms.filter(e=>_.startsWith(e.get_pathOfTerm(),parentTerm.get_pathOfTerm())); 
        }
        return [];
    }

    function getAllTermsByTermSetId(termSetId: string) {
        return createExecutionContext<SP.Taxonomy.TermCollection>((ctx, session, store, execute) => {
            let tset = store.getTermSet(new SP.Guid(termSetId));
            let terms = tset.getAllTerms();
            ctx.load(terms);
            execute(terms);
        })
        .then((terms: SP.Taxonomy.TermCollection) => {
            return terms.get_data();
        }); 
    }

    function getAllTermSetsInSiteCollectionGroup(createIfMissing:boolean = false):Promise<SP.Taxonomy.TermSet[]>{
        return createExecutionContext<SP.Taxonomy.TermSetCollection>((ctx, session, store, execute) => {
            let groups = store.getSiteCollectionGroup(ctx.get_site(),createIfMissing); 
            let terms = groups.get_termSets();
            ctx.load(terms);
            execute(terms); 
        },_spPageContextInfo.siteAbsoluteUrl)
        .then((terms: SP.Taxonomy.TermSetCollection) => {
            return terms.get_data();
        });
    }

    function getTermLabels(term:SP.Taxonomy.Term):Promise<SP.Taxonomy.Label[]>{
        return createExecutionContext<SP.Taxonomy.LabelCollection>((ctx,session,store,execute)=>{
            let labels = term.get_labels(); 
            ctx.load(labels); 
            execute(labels); 
        })
        .then((labels:SP.Taxonomy.LabelCollection)=>{
            return labels.get_data();
        }); 
    }

    function getTermLabelsById(termId:string):Promise<SP.Taxonomy.Label[]>{
        return createExecutionContext<SP.Taxonomy.LabelCollection>((ctx,session,store,execute)=>{
            let term = store.getTerm(new SP.Guid(termId)); 
            let labels = term.get_labels(); 
            ctx.load(labels); 
            execute(labels); 
        })
        .then((labels:SP.Taxonomy.LabelCollection)=>{
            return labels.get_data();
        });
    }

    function getSiteCollectionTermGroup(createIfMissing:boolean = false):Promise<SP.Taxonomy.TermGroup>{
        return createExecutionContext<SP.Taxonomy.TermGroup>((ctx,session,store,execute)=>{
            let group:SP.Taxonomy.TermGroup = store.getSiteCollectionGroup(ctx.get_site(),createIfMissing);
            ctx.load(group);
            execute(group); 
        },_spPageContextInfo.siteAbsoluteUrl);
    }

    function getLabelsForTerms(terms:SP.Taxonomy.Term[]):Promise<SP.Taxonomy.LabelCollection[]>{
        return createExecutionContext<SP.Taxonomy.LabelCollection[]>((ctx,session,store,execute)=>{
            let labels:SP.Taxonomy.LabelCollection[] = [];
            terms.forEach((term)=>{
                let l = term.get_labels(); 
                ctx.load(l); 
                labels.push(l); 
            });
            execute(labels); 
        });
    }

    function getTerms(termIds:string[]):Promise<SP.Taxonomy.Term[]>{
        return createExecutionContext<SP.Taxonomy.Term[]>((ctx,session,store,execute)=>{
            let terms:SP.Taxonomy.Term[] = [];
            termIds.forEach((id)=>{
                let t = store.getTerm(new SP.Guid(id)); 
                ctx.load(t); 
                terms.push(t); 
            });
            execute(terms); 
        });
    }

    async function getTopLevelParentOfTerm(id:string):Promise<SP.Taxonomy.Term>{
        try {
            var term = await getTermById(id);
            var path = term.get_pathOfTerm().split(';'); 
            if (path.length === 1){
                return term; 
            }else {
                path.pop(); 
                var parent:SP.Taxonomy.Term = term; 
                for(var i= 0; i<path.length;i++){
                    parent = parent.get_parent(); 
                }
                var ctx = parent.get_context(); 
                ctx.load(parent); 
                return new Promise<SP.Taxonomy.Term>((res,rej)=>{
                    ctx.executeQueryAsync(()=>{
                        res(parent); 
                    },(c,err)=>{
                        rej(err); 
                    });
                });
            }
        }catch(err){
            throw err; 
        }
    }

    async function getParentThatSatisfies(id:string,fn:(v:SP.Taxonomy.Term)=>boolean):Promise<SP.Taxonomy.Term>{
        try {
            var term = await getTermById(id);
            var path = term.get_pathOfTerm().split(';'); 
            if (path.length === 1){
                return term; 
            }else {
                path.pop(); 
                var parent:SP.Taxonomy.Term = term; 
                for(var i= 0; i<path.length;i++){
                    var parent = await getParentTermByTerm(parent); 
                    if (parent && fn(parent)){
                        return parent; 
                    }
                }
                return null; 
            }
        }catch(err){
            throw err; 
        }
    }

    
    function termSetIdFromTaxonomyField(fieldInternalName:string):Promise<string>{
        return new Promise<string>((res,rej)=>{
            let ctx = new SP.ClientContext(); 
            let web = ctx.get_web(); 
            let fields = web.get_availableFields(); 
            let field:SP.Taxonomy.TaxonomyField = fields.getByInternalNameOrTitle(fieldInternalName) as any; 
            ctx.load(field); 
            ctx.executeQueryAsync(()=>{
                res(field.get_termSetId().toString()); 
            },rej); 
        }); 
    }

    function getParentTermById(id:string){
        return createExecutionContext<SP.Taxonomy.Term>((ctx,session,store,execute)=>{
            let term = store.getTerm(new SP.Guid(id)); 
            let parent = term.get_parent(); 
            ctx.load(parent);  
            execute(parent);
        });
    }

    function getParentTermByTerm(term:SP.Taxonomy.Term){
        return new Promise<SP.Taxonomy.Term>((res,rej)=>{
            var ctx = term.get_context(); 
            var parent = term.get_parent(); 
            ctx.load(parent); 
            ctx.executeQueryAsync(()=>{
                res(parent); 
            },(c,err)=>{
                rej(err); 
            })
        }); 
    }

    function getTermById(id:string){
        return createExecutionContext<SP.Taxonomy.Term>((ctx,session,store,execute)=>{
            let term = store.getTerm(new SP.Guid(id)); 
            ctx.load(term); ; 
            execute(term);
        });
    }

    function createTerm(parentTerm:SP.Taxonomy.Term, name:string, locale:number, guid:SP.Guid, properties:IValue[]){
        return new Promise<SP.Taxonomy.Term>((res,rej)=>{
            var ctx = parentTerm.get_context(); 
            var store = parentTerm.get_termStore(); 
            let term = parentTerm.createTerm(name,locale,guid);
            _(properties).each((e)=>{
                term.setLocalCustomProperty(e.id as string,e.label);
            })
            
            store.commitAll();
            ctx.executeQueryAsync(()=>{
                res(term); 
            },(c,err)=>rej(err)); 
        });
    }

    function getTermsByIds(ids:string[]){
        return createExecutionContext<SP.Taxonomy.TermCollection>((ctx,session,store,execute)=>{
            let terms = store.getTermsById(ids.map(e=>new SP.Guid(e)));
            ctx.load(terms); ; 
            execute(terms);
        })
    }

    return {
        createTerm, 
        getTerms,
        getTermsByTermId,
        getTermsByTermSetId, 
        getAllTermsByTermSetId, 
        getTermsByIds,
        getTermById,
        getTermLabelsById,
        getTermParents,
        getParentTermByTerm,
        getTermsSubTreeFlat,
        getTopLevelParentOfTerm,
        termSetIdFromTaxonomyField,
        getParentThatSatisfies,
        getAllTermSetsInSiteCollectionGroup,
        getLabelsForTerms,
        getTermLabels,
        getSiteCollectionTermGroup,
        getParentTermById
    }; 
}

